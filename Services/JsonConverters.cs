using System.Text.Json;
using System.Text.Json.Serialization;

namespace SpiroUI.Services;

/// <summary>
/// Custom JSON converter for Sex enum to output as string
/// </summary>
public class SexJsonConverter : JsonConverter<Sex>
{
    public override Sex Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.Number)
        {
            return (Sex)reader.GetInt32();
        }
        
        var value = reader.GetString();
        return value?.ToLower() switch
        {
            "male" => Sex.Male,
            "female" => Sex.Female,
            _ => Sex.Male
        };
    }

    public override void Write(Utf8JsonWriter writer, Sex value, JsonSerializerOptions options)
    {
        writer.WriteStringValue(value.ToString());
    }
}

/// <summary>
/// Custom JSON converter to round double values to specified decimal places
/// </summary>
public class RoundedDoubleJsonConverter : JsonConverter<double>
{
    private readonly int _decimalPlaces;

    public RoundedDoubleJsonConverter(int decimalPlaces = 2)
    {
        _decimalPlaces = decimalPlaces;
    }

    public override double Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        return reader.GetDouble();
    }

    public override void Write(Utf8JsonWriter writer, double value, JsonSerializerOptions options)
    {
        if (double.IsNaN(value))
        {
            writer.WriteNullValue();
        }
        else
        {
            writer.WriteNumberValue(Math.Round(value, _decimalPlaces));
        }
    }
}

/// <summary>
/// Custom JSON converter for nullable doubles with rounding
/// </summary>
public class RoundedNullableDoubleJsonConverter : JsonConverter<double?>
{
    private readonly int _decimalPlaces;

    public RoundedNullableDoubleJsonConverter(int decimalPlaces = 2)
    {
        _decimalPlaces = decimalPlaces;
    }

    public override double? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.Null)
            return null;
        return reader.GetDouble();
    }

    public override void Write(Utf8JsonWriter writer, double? value, JsonSerializerOptions options)
    {
        if (!value.HasValue || double.IsNaN(value.Value))
        {
            writer.WriteNullValue();
        }
        else
        {
            writer.WriteNumberValue(Math.Round(value.Value, _decimalPlaces));
        }
    }
}