using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using SpiroUI.Models;
using ClosedXML.Excel;
using System.Linq;

namespace SpiroUI.Services;

public class FileProcessingService
{
    private readonly JsonSerializerOptions _pnpJsonOptions;
    private readonly JsonSerializerOptions _zakJsonOptions;

    public FileProcessingService()
    {
        // JSON options for PNP files (2 decimal places)
        _pnpJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters =
            {
                new SexJsonConverter(),
                new RoundedDoubleJsonConverter(2),
                new RoundedNullableDoubleJsonConverter(2)
            }
        };

        // JSON options for ZAK files (3 decimal places)
        _zakJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters =
            {
                new SexJsonConverter(),
                new RoundedDoubleJsonConverter(3),
                new RoundedNullableDoubleJsonConverter(3)
            }
        };
    }

    public Task<string> ParseSpirographFile(byte[] fileContent, string fileName)
    {
        var spirographData = PnpParser.ParseFile(fileContent, fileName);
        
        // Add file metadata
        var result = new
        {
            fileName = fileName,
            fileSize = fileContent.Length,
            timestamp = DateTime.UtcNow,
            data = spirographData
        };

        return Task.FromResult(JsonSerializer.Serialize(result, _pnpJsonOptions));
    }

    public Task<string> ParseZakFile(byte[] fileContent, string fileName)
    {
        // Save content to a temporary file for ZakParser
        var tempFilePath = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(tempFilePath, fileContent);
            var zakData = ZakParser.ParseFile(tempFilePath);
            
            // Fix the filename in the ZakRecord to use the original name instead of temp file name
            var correctedZakData = zakData with { FileName = fileName };
            
            // Add file metadata
            var result = new
            {
                fileName = fileName,
                fileSize = fileContent.Length,
                timestamp = DateTime.UtcNow,
                data = correctedZakData
            };

            return Task.FromResult(JsonSerializer.Serialize(result, _zakJsonOptions));
        }
        finally
        {
            if (File.Exists(tempFilePath))
                File.Delete(tempFilePath);
        }
    }

    public Task<byte[]> GenerateExcelFile(List<SpirographFile> files)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Спирография");

        // Define column headers in Russian based on PnpParser.cs comments
        int row = 1;
        int col = 1;
        var separatorColumns = new List<(int Column, string SectionName)>();
        
        // File info columns
        worksheet.Cell(row, col++).Value = "Имя файла";
        worksheet.Cell(row, col++).Value = "Размер файла";
        worksheet.Cell(row, col++).Value = "Время обработки";
        
        // Separator after file info - next section is Patient Info
        separatorColumns.Add((col, " Испытуемый "));
        worksheet.Cell(row, col++).Value = " Испытуемый ";
        
        // Patient info columns
        worksheet.Cell(row, col++).Value = "ФИО";
        worksheet.Cell(row, col++).Value = "Возраст";
        worksheet.Cell(row, col++).Value = "Вес (кг)";
        worksheet.Cell(row, col++).Value = "Рост (м)";
        worksheet.Cell(row, col++).Value = "Пол";
        worksheet.Cell(row, col++).Value = "Примечание";
        
        // Separator after patient info - next section is BTPS
        separatorColumns.Add((col, " BTPS "));
        worksheet.Cell(row, col++).Value = " BTPS ";
        
        // BTPS info
        worksheet.Cell(row, col++).Value = "BTPS коэффициент";
        worksheet.Cell(row, col++).Value = "Температура (°C)";
        worksheet.Cell(row, col++).Value = "Влажность (%)";
        worksheet.Cell(row, col++).Value = "Давление (мм рт.ст.)";
        
        // Separator after BTPS - next section is ZhEL
        separatorColumns.Add((col, " ЖЕЛ "));
        worksheet.Cell(row, col++).Value = " ЖЕЛ ";
        
        // ZhEL block columns
        worksheet.Cell(row, col++).Value = "ЕВд(л)";
        worksheet.Cell(row, col++).Value = "ЖЕЛ(л)";
        worksheet.Cell(row, col++).Value = "ДО(л)";
        worksheet.Cell(row, col++).Value = "РОвд(л)";
        worksheet.Cell(row, col++).Value = "РОвыд(л)";
        worksheet.Cell(row, col++).Value = "ДО/ЖЕЛ(%)";
        
        // Separator after ZhEL - next section is MOD
        separatorColumns.Add((col, " МОД "));
        worksheet.Cell(row, col++).Value = " МОД ";
        
        // MOD block columns
        worksheet.Cell(row, col++).Value = "ЧД(1/мин)";
        worksheet.Cell(row, col++).Value = "МОД(л/мин)";
        worksheet.Cell(row, col++).Value = "ДО МОД(л)";
        worksheet.Cell(row, col++).Value = "ПО2(мл/мин)";
        worksheet.Cell(row, col++).Value = "КИО2(мл/л)";
        worksheet.Cell(row, col++).Value = "Твыд/Твд";
        
        // Separator after MOD - next section is MVL
        separatorColumns.Add((col, " МВЛ "));
        worksheet.Cell(row, col++).Value = " МВЛ ";
        
        // MVL block columns
        worksheet.Cell(row, col++).Value = "ЧД МВЛ(1/мин)";
        worksheet.Cell(row, col++).Value = "МВЛ(л/мин)";
        worksheet.Cell(row, col++).Value = "ДО МВЛ(л)";
        worksheet.Cell(row, col++).Value = "РД(%)";
        worksheet.Cell(row, col++).Value = "МВЛ/МОД";
        
        // Separator after MVL - next section is FVC Probes
        separatorColumns.Add((col, " ФЖЕЛ Пробы "));
        worksheet.Cell(row, col++).Value = " ФЖЕЛ Пробы ";
        
        // FVC Probe columns - up to 3 probes
        for (int probeNum = 1; probeNum <= 3; probeNum++)
        {
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ФЖЕЛ(л)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ЖЕЛвд(л)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ОФВ1(л)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ОФВ1/ЖЕЛ";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ПОС(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} ОВпос(л)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} МОС25(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} МОС50(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} МОС75(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} СОС25-75(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} СОС75-85(л/с)";
            worksheet.Cell(row, col++).Value = $"Проба {probeNum} Тфжел(с)";
            
            // Separator after each probe (except the last one)
            if (probeNum < 3)
            {
                separatorColumns.Add((col, $"Проба {probeNum + 1}"));
                worksheet.Cell(row, col++).Value = $"Проба {probeNum + 1}";
            }
        }

        // Style the header row
        var headerRange = worksheet.Range(1, 1, 1, col - 1);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        
        // Make header row taller to accommodate vertical text in separator columns
        worksheet.Row(1).Height = 80;
        
        // Style separator columns
        foreach (var (separatorCol, sectionName) in separatorColumns)
        {
            var column = worksheet.Column(separatorCol);
            column.Width = 10; // Wider to accommodate section name
            column.Style.Fill.BackgroundColor = XLColor.Gray;
            column.Style.Font.Bold = true;
            column.Style.Font.FontColor = XLColor.White;
            column.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            column.Style.Alignment.TextRotation = 90; // Rotate text for better fit
        }

        // Add data rows
        row = 2;
        foreach (var file in files)
        {
            if (!string.IsNullOrEmpty(file.ParsedData))
            {
                try
                {
                    // Parse the JSON data
                    var parsedData = JsonSerializer.Deserialize<JsonElement>(file.ParsedData, _pnpJsonOptions);
                    var data = parsedData.GetProperty("data");
                    
                    col = 1;
                    
                    // File info
                    worksheet.Cell(row, col++).Value = parsedData.GetProperty("fileName").GetString();
                    worksheet.Cell(row, col++).Value = parsedData.GetProperty("fileSize").GetInt32();
                    worksheet.Cell(row, col++).Value = parsedData.GetProperty("timestamp").GetDateTime();
                    
                    // Skip separator column
                    col++;
                    
                    // Patient info
                    var patient = data.GetProperty("patient");
                    worksheet.Cell(row, col++).Value = patient.GetProperty("rawHeader").GetString();
                    worksheet.Cell(row, col++).Value = patient.GetProperty("ageYears").GetInt32();
                    worksheet.Cell(row, col++).Value = patient.GetProperty("weightKg").GetInt32();
                    worksheet.Cell(row, col++).Value = patient.GetProperty("heightM").GetDouble();
                    worksheet.Cell(row, col++).Value = patient.GetProperty("sex").GetString();
                    worksheet.Cell(row, col++).Value = patient.GetProperty("note").GetString();
                    
                    // Skip separator column
                    col++;
                    
                    // BTPS info
                    var btps = data.GetProperty("btps");
                    worksheet.Cell(row, col++).Value = btps.GetProperty("factor").GetDouble();
                    AddOptionalDouble(worksheet, row, col++, btps, "tempC");
                    AddOptionalDouble(worksheet, row, col++, btps, "humidityPct");
                    AddOptionalDouble(worksheet, row, col++, btps, "pressureMmHg");
                    
                    // Skip separator column
                    col++;
                    
                    // ZhEL block
                    if (data.TryGetProperty("zhel", out var zhel) && zhel.ValueKind != JsonValueKind.Null)
                    {
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("evd_L").GetDouble();
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("jhel_L").GetDouble();
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("do_L").GetDouble();
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("rovd_L").GetDouble();
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("rovyd_L").GetDouble();
                        worksheet.Cell(row, col++).Value = zhel.GetProperty("doOverEvd_Pct").GetDouble();
                    }
                    else
                    {
                        col += 6;
                    }
                    
                    // Skip separator column
                    col++;
                    
                    // MOD block
                    if (data.TryGetProperty("mod", out var mod) && mod.ValueKind != JsonValueKind.Null)
                    {
                        worksheet.Cell(row, col++).Value = mod.GetProperty("respiratoryRate_perMin").GetDouble();
                        worksheet.Cell(row, col++).Value = mod.GetProperty("minuteVentilation_Lpm").GetDouble();
                        worksheet.Cell(row, col++).Value = mod.GetProperty("tidalVolume_L").GetDouble();
                        // Round ПО2(мл/мин) to closest integer
                        worksheet.Cell(row, col++).Value = Math.Round(mod.GetProperty("oxygenUptake_mlMin").GetDouble());
                        // Use КИО2 вент.экв.(л/л) value for КИО2(мл/л) column
                        worksheet.Cell(row, col++).Value = mod.GetProperty("kio2VentEq_LperL").GetDouble();
                        worksheet.Cell(row, col++).Value = mod.GetProperty("texpOverTinsp").GetDouble();
                    }
                    else
                    {
                        col += 6; // Updated from 7 to 6 since we removed one column
                    }
                    
                    // Skip separator column
                    col++;
                    
                    // MVL block
                    if (data.TryGetProperty("mvl", out var mvl) && mvl.ValueKind != JsonValueKind.Null)
                    {
                        worksheet.Cell(row, col++).Value = mvl.GetProperty("respiratoryRate_perMin").GetDouble();
                        worksheet.Cell(row, col++).Value = mvl.GetProperty("maxVentilation_Lpm").GetDouble();
                        worksheet.Cell(row, col++).Value = mvl.GetProperty("tidalVolume_L").GetDouble();
                        worksheet.Cell(row, col++).Value = mvl.GetProperty("breathingReserve_Pct").GetDouble();
                        worksheet.Cell(row, col++).Value = mvl.GetProperty("mvlOverMod").GetDouble();
                    }
                    else
                    {
                        col += 5;
                    }
                    
                    // Skip separator column
                    col++;
                    
                    // FVC Probes - up to 3 probes
                    if (data.TryGetProperty("probes", out var probes))
                    {
                        var probeArray = probes.EnumerateArray().ToArray();
                        
                        for (int probeIndex = 0; probeIndex < 3; probeIndex++)
                        {
                            if (probeIndex < probeArray.Length)
                            {
                                var probe = probeArray[probeIndex];
                                var fvcUi = probe.GetProperty("fvcUi_L").GetDouble();
                                var fev1Ui = probe.GetProperty("fev1Ui_L").GetDouble();
                                
                                // Column order: ФЖЕЛ(л), ЖЕЛвд(л), ОФВ1(л), ОФВ1/ЖЕЛ, ПОС(л/с), ОВпос(л), МОС25(л/с), МОС50(л/с), МОС75(л/с), СОС25-75(л/с), СОС75-85(л/с), Тфжел(с)
                                worksheet.Cell(row, col++).Value = Math.Round(fvcUi, 2); // ФЖЕЛ(л) - use UI value, 2 decimals
                                worksheet.Cell(row, col++).Value = Math.Round(probe.GetProperty("evdUi_L").GetDouble(), 2); // ЖЕЛвд(л) - use EvdUi value, 2 decimals
                                worksheet.Cell(row, col++).Value = fev1Ui; // ОФВ1(л) - use UI value
                                if (fvcUi > 0)
                                    worksheet.Cell(row, col).Value = Math.Round(fev1Ui / fvcUi, 2); // ОФВ1/ЖЕЛ - calculate ratio, 2 decimals
                                col++;
                                worksheet.Cell(row, col++).Value = probe.GetProperty("pef_Lps").GetDouble(); // ПОС(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("ovnosUi_L").GetDouble(); // ОВпос(л) - use UI value
                                worksheet.Cell(row, col++).Value = probe.GetProperty("mos25_Lps").GetDouble(); // МОС25(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("mos50_Lps").GetDouble(); // МОС50(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("mos75_Lps").GetDouble(); // МОС75(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("sos25_75_Lps").GetDouble(); // СОС25-75(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("sos75_85_Lps").GetDouble(); // СОС75-85(л/с)
                                worksheet.Cell(row, col++).Value = probe.GetProperty("tfvc_s").GetDouble(); // Тфжел(с)
                            }
                            else
                            {
                                // Skip empty probe columns (still 12 columns)
                                col += 12;
                            }
                            
                            // Skip separator column between probes (except after the last one)
                            if (probeIndex < 2)
                            {
                                col++;
                            }
                        }
                    }
                    else
                    {
                        // Skip all probe columns if no probes (including separators)
                        col += 38; // 12 columns × 3 probes + 2 separators
                    }
                    
                    row++;
                }
                catch (Exception ex)
                {
                    // Skip files that can't be parsed
                    Console.WriteLine($"Error parsing file {file.Name}: {ex.Message}");
                }
            }
        }

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();
        
        // Save to memory stream
        using var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return Task.FromResult(memoryStream.ToArray());
    }

    private void AddOptionalDouble(IXLWorksheet worksheet, int row, int col, JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind != JsonValueKind.Null)
        {
            worksheet.Cell(row, col).Value = prop.GetDouble();
        }
    }

    public async Task<byte[]> GenerateZipArchive(List<SpirographFile> files)
    {
        using var memoryStream = new MemoryStream();
        using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
        {
            foreach (var file in files)
            {
                if (!string.IsNullOrEmpty(file.ParsedData))
                {
                    var jsonFileName = Path.GetFileNameWithoutExtension(file.Name) + ".json";
                    var entry = archive.CreateEntry(jsonFileName, CompressionLevel.Optimal);
                    
                    using var entryStream = entry.Open();
                    var jsonBytes = Encoding.UTF8.GetBytes(file.ParsedData);
                    await entryStream.WriteAsync(jsonBytes, 0, jsonBytes.Length);
                }
            }
        }

        memoryStream.Seek(0, SeekOrigin.Begin);
        return memoryStream.ToArray();
    }

    public Task<byte[]> GenerateZakExcelFile(List<SpirographFile> files)
    {
        Console.WriteLine($"GenerateZakExcelFile called with {files.Count} files");
        
        // Parse all ZAK records from JSON along with metadata
        var zakRecordsWithMetadata = new List<(ZakRecord Record, long FileSize, DateTime Timestamp)>();
        foreach (var file in files)
        {
            Console.WriteLine($"Processing file: {file.Name}, HasParsedData: {!string.IsNullOrEmpty(file.ParsedData)}");
            if (!string.IsNullOrEmpty(file.ParsedData))
            {
                try
                {
                    var parsedData = JsonSerializer.Deserialize<JsonElement>(file.ParsedData, _zakJsonOptions);
                    var dataJson = parsedData.GetProperty("data").GetRawText();
                    Console.WriteLine($"Data JSON for {file.Name}: {dataJson.Substring(0, Math.Min(500, dataJson.Length))}...");
                    
                    var zakRecord = JsonSerializer.Deserialize<ZakRecord>(dataJson, _zakJsonOptions);
                    Console.WriteLine($"Deserialized ZakRecord: {zakRecord != null}");
                    
                    if (zakRecord != null)
                    {
                        Console.WriteLine($"ZakRecord details - Section: {zakRecord.Section}, Patient null: {zakRecord.Patient == null}, Measurements null: {zakRecord.Measurements == null}");
                        
                        // Ensure all properties are never null for Excel generation
                        var safeRecord = zakRecord with
                        {
                            Patient = zakRecord.Patient ?? new PatientData(null, null, null, null, null, null, null),
                            Measurements = zakRecord.Measurements ?? new List<Measurement>(),
                            Conclusion = zakRecord.Conclusion ?? new List<ConclusionItem>(),
                            Extra = zakRecord.Extra ?? new Extras(null, null, null, null, null, null, null, null)
                        };
                        
                        var fileSize = parsedData.GetProperty("fileSize").GetInt64();
                        var timestamp = parsedData.GetProperty("timestamp").GetDateTime();
                        zakRecordsWithMetadata.Add((safeRecord, fileSize, timestamp));
                        Console.WriteLine($"Added record for {file.Name} to list. Total records: {zakRecordsWithMetadata.Count}");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to deserialize ZakRecord for {file.Name}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error parsing ZAK file {file.Name}: {ex.Message}");
                    Console.WriteLine($"ParsedData preview: {file.ParsedData?.Substring(0, Math.Min(200, file.ParsedData.Length))}");
                }
            }
        }
        
        var zakRecords = zakRecordsWithMetadata.Select(x => x.Record).ToList();
        Console.WriteLine($"Final zakRecords count: {zakRecords.Count}");
        
        if (zakRecords.Count == 0)
        {
            Console.WriteLine("No ZAK records found, creating empty workbook");
            // Still create a workbook with headers but no data
        }

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Реография");

        // Track separator columns for styling
        var separatorColumns = new List<(int Column, string SectionName)>();
        
        // Define column structure with Russian headers
        int col = 1;
        
        // File info section
        worksheet.Cell(1, col++).Value = "Имя файла";
        worksheet.Cell(1, col++).Value = "Размер файла (байт)";
        worksheet.Cell(1, col++).Value = "Время обработки";
        
        // Section separator
        separatorColumns.Add((col, " Исследование "));
        worksheet.Cell(1, col++).Value = " Исследование ";
        
        // Study info
        worksheet.Cell(1, col++).Value = "Раздел";
        worksheet.Cell(1, col++).Value = "Область";
        
        // Patient separator
        separatorColumns.Add((col, " Испытуемый "));
        worksheet.Cell(1, col++).Value = " Испытуемый ";
        
        // Patient info
        worksheet.Cell(1, col++).Value = "ФИО";
        worksheet.Cell(1, col++).Value = "Возраст";
        worksheet.Cell(1, col++).Value = "Пол";
        worksheet.Cell(1, col++).Value = "Рост (см)";
        worksheet.Cell(1, col++).Value = "Вес (кг)";
        worksheet.Cell(1, col++).Value = "Дата исследования";
        worksheet.Cell(1, col++).Value = "Комментарий";
        
        // Measurements separator - we'll add dynamic measurement columns
        separatorColumns.Add((col, " Измерения "));
        worksheet.Cell(1, col++).Value = " Измерения ";
        
        // Collect all unique measurement keys with their units to create columns
        var allMeasurementKeysWithUnits = zakRecords
            .Where(r => r.Measurements != null)
            .SelectMany(r => r.Measurements)
            .Where(m => m != null)
            .GroupBy(m => m.Key)
            .Select(g => new { 
                Key = g.Key, 
                Unit = g.FirstOrDefault(m => !string.IsNullOrEmpty(m.Unit))?.Unit ?? "" 
            })
            .OrderBy(x => x.Key)
            .ToList();
            
        // Add measurement columns (L and R for each measurement)
        foreach (var measurement in allMeasurementKeysWithUnits)
        {
            var unitSuffix = string.IsNullOrEmpty(measurement.Unit) ? "" : $" ({measurement.Unit})";
            worksheet.Cell(1, col++).Value = $"{measurement.Key} (L){unitSuffix}";
            worksheet.Cell(1, col++).Value = $"{measurement.Key} (R){unitSuffix}";
        }
        
        // Extras separator
        separatorColumns.Add((col, " Дополнительно "));
        worksheet.Cell(1, col++).Value = " Дополнительно ";
        
        // Extras columns
        worksheet.Cell(1, col++).Value = "Коэффициент асимметрии (%)";
        worksheet.Cell(1, col++).Value = "Норма асимметрии";
        worksheet.Cell(1, col++).Value = "Асимметрия кровенаполнения (кач.)";
        worksheet.Cell(1, col++).Value = "Доминирование (S/D)";
        worksheet.Cell(1, col++).Value = "Доминирующая сторона";
        worksheet.Cell(1, col++).Value = "ЧСС (уд/мин)";
        worksheet.Cell(1, col++).Value = "ЧСС мин (уд/мин)";
        worksheet.Cell(1, col++).Value = "ЧСС макс (уд/мин)";
        
        // Conclusions separator
        separatorColumns.Add((col, " Заключения "));
        worksheet.Cell(1, col++).Value = " Заключения ";
        
        // Collect all unique conclusion keys
        var allConclusionKeys = zakRecords
            .Where(r => r.Conclusion != null)
            .SelectMany(r => r.Conclusion)
            .Where(c => c != null)
            .Select(c => c.Key)
            .Distinct()
            .OrderBy(k => k)
            .ToList();
            
        // Add conclusion columns
        foreach (var key in allConclusionKeys)
        {
            worksheet.Cell(1, col++).Value = $"{key} - Значение";
            worksheet.Cell(1, col++).Value = $"{key} - Примечание";
            worksheet.Cell(1, col++).Value = $"{key} - Отклонение (%)";
            worksheet.Cell(1, col++).Value = $"{key} - Направление";
        }

        // Style the header row
        var headerRange = worksheet.Range(1, 1, 1, col - 1);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        
        // Make header row taller to accommodate vertical text in separator columns
        worksheet.Row(1).Height = 80;
        headerRange.Style.Alignment.WrapText = true;
        
        // Style separator columns
        foreach (var (separatorCol, sectionName) in separatorColumns)
        {
            var column = worksheet.Column(separatorCol);
            column.Width = 3;
            column.Style.Fill.BackgroundColor = XLColor.Gray;
            column.Style.Font.Bold = true;
            column.Style.Font.FontColor = XLColor.White;
            column.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            column.Style.Alignment.TextRotation = 90;
        }
        
        // Add data rows
        int row = 2;
        for (int i = 0; i < zakRecordsWithMetadata.Count; i++)
        {
            var (record, fileSize, timestamp) = zakRecordsWithMetadata[i];
            col = 1;
            
            // File info
            worksheet.Cell(row, col++).Value = record.FileName;
            worksheet.Cell(row, col++).Value = fileSize;
            worksheet.Cell(row, col++).Value = timestamp;
            
            // Skip separator
            col++;
            
            // Study info
            worksheet.Cell(row, col++).Value = record.Section;
            worksheet.Cell(row, col++).Value = record.Area ?? "";
            
            // Skip patient separator
            col++;
            
            // Patient info
            worksheet.Cell(row, col++).Value = record.Patient?.FullName ?? "";
            if (record.Patient?.Age.HasValue == true)
                worksheet.Cell(row, col).Value = record.Patient.Age.Value;
            col++;
            
            worksheet.Cell(row, col++).Value = record.Patient?.Sex ?? "";
            
            if (record.Patient?.Height.HasValue == true)
                worksheet.Cell(row, col).Value = record.Patient.Height.Value;
            col++;
            
            if (record.Patient?.Weight.HasValue == true)
                worksheet.Cell(row, col).Value = record.Patient.Weight.Value;
            col++;
            if (record.Patient?.Date.HasValue == true)
                worksheet.Cell(row, col++).Value = record.Patient.Date.Value;
            else
                col++;
            worksheet.Cell(row, col++).Value = record.Patient?.Comment ?? "";
            
            // Skip measurements separator
            col++;
            
            // Measurements - populate based on the predefined columns
            foreach (var measurement in allMeasurementKeysWithUnits)
            {
                var leftMeasurement = record.Measurements?.FirstOrDefault(m => m.Key == measurement.Key && m.Side == "L");
                var rightMeasurement = record.Measurements?.FirstOrDefault(m => m.Key == measurement.Key && m.Side == "R");
                
                if (leftMeasurement?.Value.HasValue == true)
                    worksheet.Cell(row, col).Value = Math.Round(leftMeasurement.Value.Value, 3);
                col++;
                
                if (rightMeasurement?.Value.HasValue == true)
                    worksheet.Cell(row, col).Value = Math.Round(rightMeasurement.Value.Value, 3);
                col++;
            }
            
            // Skip extras separator
            col++;
            
            // Extras
            if (record.Extra?.AsymmetryCoeffPercent.HasValue == true)
                worksheet.Cell(row, col).Value = Math.Round(record.Extra.AsymmetryCoeffPercent.Value, 3);
            col++;
            worksheet.Cell(row, col++).Value = record.Extra?.AsymmetryNormText ?? "";
            worksheet.Cell(row, col++).Value = record.Extra?.AsymmetryQualitative ?? "";
            worksheet.Cell(row, col++).Value = record.Extra?.AsymmetryDominanceCode ?? "";
            worksheet.Cell(row, col++).Value = record.Extra?.AsymmetryDominanceSide ?? "";
            if (record.Extra?.HeartRateBpm.HasValue == true)
                worksheet.Cell(row, col).Value = record.Extra.HeartRateBpm.Value;
            col++;
            
            if (record.Extra?.HeartRateLow.HasValue == true)
                worksheet.Cell(row, col).Value = record.Extra.HeartRateLow.Value;
            col++;
            
            if (record.Extra?.HeartRateHigh.HasValue == true)
                worksheet.Cell(row, col).Value = record.Extra.HeartRateHigh.Value;
            col++;
            
            // Skip conclusions separator
            col++;
            
            // Conclusions
            foreach (var key in allConclusionKeys)
            {
                var conclusion = record.Conclusion?.FirstOrDefault(c => c.Key == key);
                worksheet.Cell(row, col++).Value = conclusion?.Value ?? "";
                worksheet.Cell(row, col++).Value = conclusion?.Note ?? "";
                if (conclusion?.DeltaPercent.HasValue == true)
                    worksheet.Cell(row, col).Value = Math.Round(conclusion.DeltaPercent.Value, 3);
                col++;
                worksheet.Cell(row, col++).Value = conclusion?.DeltaDirection ?? "";
            }
            
            row++;
        }
        
        // Auto-fit columns
        worksheet.Columns().AdjustToContents();
        
        // Save to memory stream
        using var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return Task.FromResult(memoryStream.ToArray());
    }
}