// ------------------------------------------------------------------------------------
//  Pulmo-4 PNP File Reader (single-file C#, .NET 8+) – clean, production-ready version
// ------------------------------------------------------------------------------------
//  Decodes diagnostic data saved by the Pulmo-4 spirometer (.pnp files).
//  Extracts:
//      • Demographics (name, age, sex, height, weight)
//      • Automatic BTPS factor (with fallback constant)
//      • Resting vital capacity block (ZhEL)
//      • Minute ventilation block (MOD)  + vendor-exact Texp/Tinsp (ΣTexp / ΣTinsp)
//      • Maximum voluntary ventilation block (MVL)
//      • Up-to-three forced-vital-capacity probes (FZhEL, FZhE1, FZhE2)
//
//  Usage example:
//      var rec = Pulmo4.PnpParser.ParseFile("2-1.pnp");
//      Console.WriteLine($"Tвыд/Твд: {rec.Mod!.TexpOverTinsp:F2}");   // 1.30
// ------------------------------------------------------------------------------------

using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace SpiroUI.Services;

#region ── Domain models ─────────────────────────────────────────────────────────────

public enum Sex
{
    Male = 0,
    Female = 1
}

public sealed record BtpsInfo(
    bool FoundInFile,
    double Factor,
    double? TempC = null,
    double? HumidityPct = null,
    double? PressureMmHg = null);

public sealed record Demographics(
    string RawHeader,
    int AgeYears,
    int WeightKg,
    double HeightM,
    Sex Sex,
    string Note);

public sealed record FvcProbe(
    int Index,
    double Fvc_L,
    double Fev1_L,
    double Evd_L,
    double Pef_Lps,
    double Ovnos_L,
    double Mos25_Lps,
    double Mos50_Lps,
    double Mos75_Lps,
    double Sos25_75_Lps,
    double Sos75_85_Lps,
    double Tfvc_s,
    double FvcUi_L,
    double Fev1Ui_L,
    double OvnosUi_L,
    double EvdUi_L);

public sealed record ZhelBlock(
    double Evd_L,
    double Jhel_L,
    double Do_L,
    double Rovd_L,
    double Rovyd_L,
    double DoOverEvd_Pct);

public sealed record ModBlock(
    double RespiratoryRate_perMin,
    double MinuteVentilation_Lpm,
    double TidalVolume_L,
    double OxygenUptake_mlMin,
    double Kio2_mlPerL, // kO2: мл O2 на 1 литр вентиляции
    double Kio2VentEq_LperL, // вентиляционный эквивалент по O2: л воздуха на 1 л O2
    double TexpOverTinsp);

public sealed record MvlBlock(
    double RespiratoryRate_perMin,
    double MaxVentilation_Lpm,
    double TidalVolume_L,
    double BreathingReserve_Pct, // РД(%)
    double MvlOverMod); // МВЛ/МОД (безразмерное отношение)

public sealed record PnpRecord(
    string FileName,
    // PredictedSet Predicted,
    Demographics Patient,
    BtpsInfo Btps,
    ZhelBlock? Zhel,
    ModBlock? Mod,
    MvlBlock? Mvl,
    IReadOnlyList<FvcProbe> Probes);

#endregion

// ================================================================================
//  Parser – single public class
// ================================================================================
public static class PnpParser
{
    #region ── Binary tag constants ───────────────────────────────────────────────

    private static readonly byte[] TagMod = Encoding.ASCII.GetBytes("MOD    ");
    private static readonly byte[] TagMvl = Encoding.ASCII.GetBytes("MVL    ");
    private static readonly byte[] TagZhel = Encoding.ASCII.GetBytes("* ZhEL *");
    private static readonly byte[] TagF1 = Encoding.ASCII.GetBytes("* FZhEL* ");
    private static readonly byte[] TagF2 = Encoding.ASCII.GetBytes("* FZhE1* ");
    private static readonly byte[] TagF3 = Encoding.ASCII.GetBytes("* FZhE2* ");
    private static readonly byte[][] AllTags = { TagMod, TagMvl, TagZhel, TagF1, TagF2, TagF3 };

    #endregion

    // --------------------------------------------------------------------------------
    public static PnpRecord ParseFile(string filePath,
        double defaultBtpsFactor = 1.081,
        double kio2_mlPerL = 25.0)
    {
        var fileBytes = File.ReadAllBytes(filePath);
        var fileName = Path.GetFileName(filePath);
        return Parse(fileBytes, fileName, defaultBtpsFactor, kio2_mlPerL);
    }

    public static PnpRecord ParseFile(byte[] fileContent, string fileName,
        double defaultBtpsFactor = 1.081,
        double kio2_mlPerL = 25.0)
    {
        return Parse(fileContent, fileName, defaultBtpsFactor, kio2_mlPerL);
    }

    // --------------------------------------------------------------------------------
    private static PnpRecord Parse(ReadOnlySpan<byte> buffer,
        string fileName,
        double defaultBtpsFactor,
        double kio2_mlPerL)
    {
        // 1) Demographics & BTPS
        var patient = ExtractDemographics(buffer);
        var btps = ExtractBtps(buffer, defaultBtpsFactor);

        // Containers for parsed blocks
        var probes = new List<FvcProbe>();
        ZhelBlock? zhel = null;
        ModBlock? mod = null;
        MvlBlock? mvl = null;

        // 2) Forced-vital-capacity probes
        ParseFvcProbe(buffer, TagF1, 1, probes, btps);
        ParseFvcProbe(buffer, TagF2, 2, probes, btps);
        ParseFvcProbe(buffer, TagF3, 3, probes, btps);

        // 3) ZhEL block (resting vital capacity)
        var zhelPos = buffer.IndexOf(TagZhel);
        if (zhelPos >= 0 && zhelPos + TagZhel.Length + 20 <= buffer.Length)
        {
            var values = ReadFloatArray(buffer.Slice(zhelPos + TagZhel.Length, 20));
            zhel = new ZhelBlock(
                values[0] * 1e-3,
                values[1] * 1e-3,
                values[2] * 1e-3,
                values[3] * 1e-3,
                values[4] * 1e-3,
                values[1] > 0 ? 100.0 * values[2] / values[1] : double.NaN
            );
        }

        // 4) MOD block
        var modPos = buffer.IndexOf(TagMod);
        if (modPos >= 0 && modPos + TagMod.Length + 12 <= buffer.Length)
        {
            var payloadStart = modPos + TagMod.Length;
            var modVals = ReadFloatArray(buffer.Slice(payloadStart, 12));
            var volumeCurve = ReadInt16Block(buffer, payloadStart + 12);

            double kio2 = kio2_mlPerL; // уже есть параметр метода Parse(...)
            double vo2 = modVals[1] * kio2; // VO2 = VE * kO2
            double kio2VentEq = kio2 > 0 ? 1000.0 / kio2 : double.NaN; // КИО2 (л/л) = 1000 / (мл/л)

            var texpTinsp = CalculateTexpOverTinsp(volumeCurve);
            mod = new ModBlock(
                modVals[0],
                modVals[1],
                modVals[2],
                vo2,
                Kio2_mlPerL: kio2,
                Kio2VentEq_LperL: kio2VentEq,
                texpTinsp);
        }

        // 5) MVL block
        int mvlPos = buffer.IndexOf(TagMvl);
        if (mvlPos >= 0 && mvlPos + TagMvl.Length + 12 <= buffer.Length)
        {
            float[] mvlVals = ReadFloatArray(buffer.Slice(mvlPos + TagMvl.Length, 12));
            double rr = mvlVals[0];
            double mvlLpm = mvlVals[1];
            double vtL = mvlVals[2];

            // РД(%) = 100 * (1 - МОД/МВЛ). МОД берём из уже распарсенного блока MOD.
            double rdPct =
                (mod is { MinuteVentilation_Lpm: > 0 } && mvlLpm > 0)
                    ? 100.0 * (1.0 - mod!.MinuteVentilation_Lpm / mvlLpm)
                    : double.NaN;

            // МВЛ/МОД
            double mvlOverMod =
                (mod is { MinuteVentilation_Lpm: > 0 })
                    ? mvlLpm / mod!.MinuteVentilation_Lpm
                    : double.NaN;

            mvl = new MvlBlock(rr, mvlLpm, vtL, rdPct, mvlOverMod);
        }

        return new PnpRecord(fileName, patient, btps, zhel, mod, mvl, probes);
    }

    // ───── local: probe parser (static – avoids capturing ref structs) ───────────
    private static void ParseFvcProbe(ReadOnlySpan<byte> buf, byte[] tag, int probeIndex,
        List<FvcProbe> probes, BtpsInfo btps)
    {
        var tagPos = buf.IndexOf(tag);
        if (tagPos < 0) return;
        var payloadStart = tagPos + tag.Length;
        if (payloadStart + 48 > buf.Length) return;
        var values = ReadFloatArray(buf.Slice(payloadStart, 48));

        // Field indices in 12-float structure
        const int IdxFvc = 0;
        const int IdxEvd = 1;
        const int IdxFev1 = 2;
        const int IdxPef = 4;
        const int IdxOvnos = 5;
        const int IdxMos25 = 6;
        const int IdxMos50 = 7;
        const int IdxMos75 = 8;
        const int IdxSos25_75 = 9;
        const int IdxSos75_85 = 10;
        const int IdxTfvc = 11;

        var btpsFactor = btps.Factor;
        var rawFvc = values[IdxFvc] * 1e-3;
        var rawFev1 = values[IdxFev1] * 1e-3;
        var rawEvd = values[IdxEvd] * 1e-3;
        var rawOv = values[IdxOvnos] * 1e-3;

        probes.Add(new FvcProbe(
            probeIndex,
            rawFvc,
            rawFev1,
            rawEvd * btpsFactor,
            values[IdxPef] * 1e-3,
            rawOv,
            values[IdxMos25] * 1e-3,
            values[IdxMos50] * 1e-3,
            values[IdxMos75] * 1e-3,
            values[IdxSos25_75],
            values[IdxSos75_85],
            values[IdxTfvc] * 0.01,
            rawFvc * btpsFactor,
            rawFev1 * btpsFactor,
            rawOv * btpsFactor,
            rawEvd)
        );
    }

    #region ── Byte-level helpers ─────────────────────────────────────────────────

    private static float[] ReadFloatArray(ReadOnlySpan<byte> bytes)
    {
        var result = new float[bytes.Length / 4];
        for (var i = 0; i < result.Length; i++)
            result[i] = BitConverter.ToSingle(bytes.Slice(4 * i, 4));
        return result;
    }

    // Returns a slice with the Int16 MOD volume curve until the next tag or EOF
    private static ReadOnlySpan<short> ReadInt16Block(ReadOnlySpan<byte> src, int start)
    {
        var end = src.Length;

        foreach (var tag in AllTags)
        {
            // ищем тег в хвосте, начиная с start
            var rel = src.Slice(start).IndexOf(tag); // IndexOf(ReadOnlySpan<T>) – OK
            if (rel < 0) continue;

            var pos = start + rel; // абсолютная позиция тега
            if (pos < end) end = pos;
        }

        var count = (end - start) / 2; // кол-во Int16
        return MemoryMarshal.Cast<byte, short>(src.Slice(start, count * 2));
    }

    #endregion

    #region ── Physiology helpers ───────────────────────────────────────────────

    private static double CalculateTexpOverTinsp(ReadOnlySpan<short> volumeCurve)
    {
        if (volumeCurve.Length > 2)
            return (double)volumeCurve[0] / volumeCurve[1];

        return Double.NaN;
    }

    #endregion

    #region ── Demographics & BTPS extraction ───────────────────────────────────

    private static Demographics ExtractDemographics(ReadOnlySpan<byte> src)
    {
        var (nameFromHdr, noteFromHdr) = ReadHeaderNameAndNote(src);

        // сканируем каждый байт, иначе можно промахнуться по фактическому смещению структуры
        int maxOff = Math.Min(4096, src.Length - 9); // -9: нам нужно читать off..off+8
        for (int off = 0; off < maxOff; off++) // ВАЖНО: off++ (не off += 2)
        {
            ushort age = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off, 2));
            ushort kg = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off + 2, 2));
            float height = BitConverter.ToSingle(src.Slice(off + 4, 4));
            byte sex = src[off + 8];

            bool plausible =
                age is >= 5 and <= 120 &&
                kg is >= 20 and <= 200 &&
                height is >= 1.2f and <= 2.5f &&
                (sex == 0 || sex == 1);

            if (plausible)
                return new Demographics(nameFromHdr, age, kg, Math.Round(height, 3), (Sex)sex, noteFromHdr);
        }

        // если не нашли — возвращаем header и значения по умолчанию
        return new Demographics(nameFromHdr, 0, 0, 0, Sex.Male, string.Empty);
    }

    private static string ReadHeaderString(ReadOnlySpan<byte> src)
    {
        // Возвращаем только имя (для обратной совместимости)
        var (name, _) = ReadHeaderNameAndNote(src);
        return name;
    }

    /// <summary>
    /// Парсит первые байты файла как CP1251-строку заголовка и возвращает (Name, Note).
    /// Формат: [служебн]<ФИО>$<ПРИМЕЧАНИЕ>\0 …
    /// После извлечения принудительно отбрасываем последний символ примечания (если он есть).
    /// </summary>
    private static (string Name, string Note) ReadHeaderNameAndNote(ReadOnlySpan<byte> src)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var enc1251 = Encoding.GetEncoding(1251);

        // Декодируем первые <=256 байт (ANSI однобайтовая, безопасно)
        string s = enc1251.GetString(src[..Math.Min(256, src.Length)]);

        // 1) Режем по первому служебному разделителю (NULL, RS/US/GS)
        int cutCtrl = s.IndexOfAny(new[] { '\0', '\x1E', '\x1F', '\x1D' });
        if (cutCtrl >= 0) s = s[..cutCtrl];

        // 2) Удаляем управляющие (<0x20), оставляя '$' как разделитель
        if (s.Length > 0)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (ch == '$' || ch >= ' ')
                    sb.Append(ch);
            }

            s = sb.ToString();
        }

        s = s.Trim();

        // 3) Делим на имя и примечание
        int p = s.IndexOf('$');
        string name, note;
        if (p < 0)
        {
            name = s;
            note = string.Empty;
        }
        else
        {
            name = s[..p].Trim();
            note = s[(p + 1)..].Trim();

            // КЛЮЧЕВАЯ ПРАВКА: всегда отбрасываем последний символ примечания, если он есть
            note = DropLastIfAny(note);
        }

        // Финальная чистка имени — только печатаемые «именные» символы
        name = CleanupName(name);

        return (name, note);
    }

/* ====================== Хелперы ====================== */

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    static string DropLastIfAny(string s) => string.IsNullOrEmpty(s) ? s : s[..(s.Length - 1)];

    static string CleanupName(string s)
    {
        s = s.Trim();
        if (s.Length == 0) return s;

        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            if (char.IsLetter(ch) || ch == ' ' || ch == '-' || ch == '.' || ch == '\'')
                sb.Append(ch);
        }

        return sb.ToString().Trim();
    }

    private static BtpsInfo ExtractBtps(ReadOnlySpan<byte> src, double defaultFactor)
    {
        for (var off = 0; off + 6 <= src.Length; off += 2)
        {
            var t10 = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off, 2)); // °C ×10
            var rh10 = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off + 2, 2)); // %RH ×10
            var p = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off + 4, 2)); // mmHg

            var plausible = t10 is >= 150 and <= 350 && p is >= 650 and <= 820;
            if (!plausible) continue;

            var T = t10 / 10.0;
            var RH = Math.Min(rh10 / 10.0, 100.0);
            double P = p;

            if (P <= 47.0) return new BtpsInfo(false, defaultFactor); // защита от деления на ~0

            var factor = ComputeBtpsFactor(T, /*RH,*/ P);
            return new BtpsInfo(true, factor, T, RH, P);
        }

        return new BtpsInfo(false, defaultFactor);
    }

    /*private static double ComputeBtpsFactor(double tempC, double rhPct, double pressureMmHg)
    {
        // Формула Tetens для насыщенного пара, hPa
        var es_hPa = 6.1078 * Math.Pow(10.0, 7.5 * tempC / (237.3 + tempC));
        var es_mmHg = es_hPa * 0.750061683;
        var ph2oAmb = es_mmHg * rhPct / 100.0;

        const double Ph2oBody = 47.0; // при 37 °C, насыщенный
        var kelvinBody = 310.0;
        var kelvinAmbient = 273.15 + tempC;

        return (pressureMmHg - ph2oAmb) / (pressureMmHg - Ph2oBody) *
               (kelvinBody / kelvinAmbient);
    }*/

    // Насыщенное давление пара воды при целой температуре T (°C), мм рт. ст.
    // Таблица сгенерирована по формуле Buck (over water) и округлена до 2 знаков (AwayFromZero).
    // Индекс: 0..60 °C. В UI используются 10..35 °C.
    private static readonly double[] PH2O_MmHg_IntC = new double[] {
        4.58, 4.93, 5.29, 5.69, 6.10, 6.54, 7.01, 7.51, 8.05, 8.61,
        9.21, 9.85, 10.52, 11.23, 11.99, 12.79, 13.64, 14.53, 15.48, 16.48,
        17.54, 18.66, 19.83, 21.08, 22.39, 23.77, 25.22, 26.75, 28.36, 30.06,
        31.84, 33.72, 35.68, 37.75, 39.92, 42.20, 44.60, 47.10, 49.73, 52.49,
        55.37, 58.39, 61.56, 64.87, 68.33, 71.95, 75.73, 79.69, 83.81, 88.13,
        92.63, 97.32, 102.22, 107.33, 112.66, 118.21, 123.99, 130.01, 136.28, 142.81,
        149.60
    };

    private static double SaturationVaporPressureMmHgAtC_Int(int tC)
    {
        if (tC < 0) tC = 0;
        if (tC > 60) tC = 60;
        return PH2O_MmHg_IntC[tC];
    }

    /// <summary>
    /// BTPS множитель «как в старой программе»:
    /// BTPS = ((P - PH2O(T_saturated)) / (P - 47)) * (310 / (273 + T_int))
    /// Где:
    ///   - T_int — целая температура °C (усечение),
    ///   - PH2O(T) — насыщенный пар при T_int из таблицы,
    ///   - RH игнорируется.
    /// </summary>
    public static double ComputeBtpsFactor(double tempC, double pressureMmHg)
    {
        int tIntC = (int)tempC;          // усечение до целых градусов
        if (tIntC < 10) tIntC = 10;      // UI 10..35 °C
        if (tIntC > 35) tIntC = 35;

        double pH2O_Amb = SaturationVaporPressureMmHgAtC_Int(tIntC);

        const double PH2O_Body = 47.0;   // мм рт. ст. при 37°C
        const double K_Body = 310.0;     // как в оригинале
        double k_Amb = 273.0 + tIntC;    // без 0.15

        if (pressureMmHg <= PH2O_Body) return 1.0;

        double num = pressureMmHg - pH2O_Amb;
        double den = pressureMmHg - PH2O_Body;
        return (num / den) * (K_Body / k_Amb);
    }
    
    #endregion
} // class PnpParser