using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using SpiroUI.Models;

namespace SpiroUI.Services;

public class FileProcessingService
{
    private readonly JsonSerializerOptions _jsonOptions;

    public FileProcessingService()
    {
        _jsonOptions = new JsonSerializerOptions
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

        return Task.FromResult(JsonSerializer.Serialize(result, _jsonOptions));
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
}