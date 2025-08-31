using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;

var builder = WebApplication.CreateBuilder(args);

// Configuración de CORS
var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
        builder => builder.WithOrigins(
                "https://sismed-dirisle.googlesites.cloud",
                "https://luiscarrillo7.github.io",
                "https://sismed-frontend.onrender.com",
                "https://consulta-sismed.googlesites.cloud",
                "https://sismed-frontend.pages.dev",
                "https://sismed-frontend.vercel.app")
            .AllowAnyHeader()
            .AllowAnyMethod());
});

var app = builder.Build();
app.UseCors(MyAllowSpecificOrigins);

static async Task<SheetsService> GetSheetsService(string base64Json)
{
    var jsonString = Encoding.UTF8.GetString(Convert.FromBase64String(base64Json));
    var credential = GoogleCredential.FromJson(jsonString)
        .CreateScoped(SheetsService.Scope.SpreadsheetsReadonly);

    return await Task.FromResult(new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "MinimalApiSheets"
    }));
}

app.MapGet("/leer-sheet", async (string id) =>
{
    try
    {
        var base64Json = Environment.GetEnvironmentVariable("GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_BASE64");
        var sheetName = Environment.GetEnvironmentVariable("GOOGLE_SHEET_NAME") ?? "Hoja1";

        if (string.IsNullOrEmpty(base64Json))
            return Results.Problem("❌ Falta la variable de entorno GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_BASE64.");

        var medicamentosSheetId = Environment.GetEnvironmentVariable("medicamentos");
        if (string.IsNullOrEmpty(medicamentosSheetId))
            return Results.Problem("❌ Falta la variable de entorno GOOGLE_SHEET_MEDICAMENTOS_ID.");

        var sheetsService = await GetSheetsService(base64Json);

        // 1. Cargar datos de medicamentos
        var medicamentosRange = "Sheet1!A:C";
        var medicamentosRequest = sheetsService.Spreadsheets.Values.Get(medicamentosSheetId, medicamentosRange);
        var medicamentosResponse = await medicamentosRequest.ExecuteAsync();

        var medicamentosLookup = new Dictionary<string, (string nombre, string datoC)>();
        if (medicamentosResponse.Values?.Count > 0)
        {
            foreach (var row in medicamentosResponse.Values.Skip(1))
            {
                if (row?.Count >= 1 && !string.IsNullOrEmpty(row[0]?.ToString()))
                {
                    var codigo = row[0].ToString()?.Trim() ?? string.Empty;
                    var nombre = row.Count > 1 ? row[1]?.ToString() : null;
                    var datoC = row.Count > 2 ? row[2]?.ToString() : null;

                    if (!string.IsNullOrEmpty(codigo) && !medicamentosLookup.ContainsKey(codigo))
                    {
                        medicamentosLookup.Add(codigo,
                            (nombre ?? "Nombre Desconocido",
                             datoC ?? "S/P"));
                    }
                }
            }
        }

        // 2. Obtener hoja principal
        string spreadsheetId;
        string responseMessage;

        if (id == "1")
        {
            spreadsheetId = Environment.GetEnvironmentVariable("id_farmaminsaelagustino") ?? string.Empty;
            if (string.IsNullOrEmpty(spreadsheetId))
                return Results.Problem("❌ La variable de entorno GOOGLE_SHEET_ID1 no está configurada.");
            responseMessage = $"✅ Datos leídos correctamente del Google Sheet con ID '{id}'.";
        }
        else if (id == "2")
        {
            spreadsheetId = Environment.GetEnvironmentVariable("id_famaminsaate") ?? string.Empty;
            if (string.IsNullOrEmpty(spreadsheetId))
                return Results.Problem("❌ La variable de entorno GOOGLE_SHEET_ID2 no está configurada.");
            responseMessage = $"✅ Datos leídos correctamente del Google Sheet con ID '{id}'.";
        }
        else if (id == "3")
        {
            spreadsheetId = Environment.GetEnvironmentVariable("id_farmaminsachosica") ?? string.Empty;
            if (string.IsNullOrEmpty(spreadsheetId))
                return Results.Problem("❌ La variable de entorno SHEET_FARMAMINSA_ATE no está configurada.");
            responseMessage = "✅ Datos leídos correctamente de FARMAMINSA ATE.";
        }
        else
        {
            return Results.Problem("❌ ID de hoja inválido. Use '1', '2' o 'farma' como parámetro.");
        }

        // 3. Leer datos principales
        var range = $"{sheetName}!A:ZZ";
        var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
        var response = await request.ExecuteAsync();

        var processedValues = new List<object[]>();
        if (response.Values?.Count > 0)
        {
            // ENCABEZADOS
            processedValues.Add(new object[] {
                "MEDICAMENTO_NOMBRE",
                "PRESENTACION",
                "FECHA_STOCK",
                "STOCK",
                "STKPRECIO"
            });

            const int MEDCOD_INDEX = 1;
            const int STKSALDO_INDEX = 2;
            const int STKPRECIO_INDEX = 3;
            const int STKFECHULT_INDEX = 4;

            foreach (var row in response.Values.Skip(1))
            {
                var medCodRaw = row?.Count > MEDCOD_INDEX ? row[MEDCOD_INDEX]?.ToString()?.Trim() : string.Empty;
                var stkSaldoValue = row?.Count > STKSALDO_INDEX ? row[STKSALDO_INDEX]?.ToString() : "N/A";
                var stkFechUltValue = row?.Count > STKFECHULT_INDEX ? row[STKFECHULT_INDEX]?.ToString() : "N/A";

                // Formatear STKPRECIO
                string stkPrecioValue = "S/ 0.00";
                if (row?.Count > STKPRECIO_INDEX)
                {
                    var rawPrecio = row[STKPRECIO_INDEX]?.ToString();
                    if (double.TryParse(rawPrecio, out double precio))
                    {
                        stkPrecioValue = $"S/ {precio:0.00}";
                    }
                }

                var (medicamentoNombre, datoCValue) = ("Desconocido", "N/A");

                if (!string.IsNullOrEmpty(medCodRaw) && medicamentosLookup.TryGetValue(medCodRaw, out var medicamentoInfo))
                {
                    medicamentoNombre = medicamentoInfo.nombre;
                    datoCValue = medicamentoInfo.datoC;
                }

                processedValues.Add(new object[] {
                    medicamentoNombre,
                    datoCValue,
                    stkFechUltValue ?? "N/A",
                    stkSaldoValue ?? "N/A",
                    stkPrecioValue
                });
            }
        }

        return Results.Json(new
        {
            message = responseMessage,
            valores = processedValues
        });
    }
    catch (Google.GoogleApiException ex)
    {
        return Results.Problem($"❌ Error de Google API: {ex.Message}");
    }
    catch (Exception ex)
    {
        return Results.Problem($"❌ Error inesperado: {ex.Message}");
    }
});

app.Run();