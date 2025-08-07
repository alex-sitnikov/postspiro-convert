using Microsoft.AspNetCore.Components.Forms;
using System.Text.Json.Serialization;

namespace SpiroUI.Models;

public class SpirographFile
{
    public string Name { get; set; } = string.Empty;
    public long Size { get; set; }
    
    [JsonIgnore]
    public IBrowserFile? BrowserFile { get; set; }
    
    [JsonIgnore]
    public byte[]? Content { get; set; }
    
    public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;
    public string? ErrorMessage { get; set; }
    public string? ParsedData { get; set; } // JSON string
}