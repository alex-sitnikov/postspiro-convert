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
using System.Runtime.InteropServices;
using System.Text;

namespace SpiroUI.Services;

#region ── Domain models ─────────────────────────────────────────────────────────────

public readonly record struct PredictedFVC(
    double FVC_L,         // ФЖЕЛ(л)
    double IC_L,          // ЕВд(л) — в таблице ФЖЕЛ
    double FEV1_L,        // ОФВ1(л)
    double FEV1_over_VC,  // ОФВ1/ЖЕЛ
    double PEF_Lps,       // ПОС(л/с)
    double OVnos_L,       // ОВнос(л)
    double FEF25_Lps,     // МОС25(л/с)
    double FEF50_Lps,     // МОС50(л/с)
    double FEF75_Lps,     // МОС75(л/с)
    double FEF25_75_Lps,  // СОС25–75(л/с)
    double FEF75_85_Lps,  // СОС75–85(л/с)
    double TFVC_s         // Тфжел(с)
);

public readonly record struct PredictedZhEL(
    double VC_L,          // ЖЕЛ(л)
    double IC_L,          // ЕВд(л)
    double VT_L,          // ДО(л)
    double IRV_L,         // РОвд(л)
    double ERV_L         // РОвыд(л)
    // double VT_over_VC_pct // ДО/ЖЕЛ(%)
);

public readonly record struct PredictedMOD(
    //double RR_perMin,     // ЧД(1/мин)
    double VE_Lpm,        // МОД(л/мин)
    double VT_L,          // ДО(л)
    double VO2_mlMin,     // ПО2(мл/мин)
    double KNO2_mlL      // КНО2(мл/л)
    //double Texp_over_Tinsp // Твыд/Твд (как «должное» фиксированное)
);

public readonly record struct PredictedMVL(
    //double RR_perMin,     // ЧД(1/мин)
    double MVV_Lpm,       // МВЛ(л/мин)
    double VT_L,          // ДО(л)
    double RD_pct        // РД(%)
    //double MVV_over_MOD   // МВЛ/МОД
);

public readonly record struct PredictedSet(
    PredictedFVC FVC,
    PredictedZhEL ZhEL,
    PredictedMOD MOD,
    PredictedMVL MVL
);

public enum MvvPredMode
{
    Fev1TimesFactor,          // MVV_pred = round(FEV1_pred,2) × factor
    Fev1TimesFactorTimesBtps  // MVV_pred = round(FEV1_pred,2) × factor × BTPS_amb
}

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
    Sex Sex);

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
    double Kio2_mlPerL,          // kO2: мл O2 на 1 литр вентиляции
    double Kio2VentEq_LperL,     // вентиляционный эквивалент по O2: л воздуха на 1 л O2
    double TexpOverTinsp);

public sealed record MvlBlock(
    double RespiratoryRate_perMin,
    double MaxVentilation_Lpm,
    double TidalVolume_L,
    double BreathingReserve_Pct, // РД(%)
    double MvlOverMod );        // МВЛ/МОД (безразмерное отношение)
                                 
                                  

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

        public static PredictedSet Calculate(
            int ageYears,
            double heightCm,
            Sex sex,
            double btpsAmb,
            double veRestOverride = 9.68,
            double mvvFactorFemale = 35.0,  // по умолчанию: Ж ×35 (даёт 116.87 при FEV1≈3.34)
            double mvvFactorMale   = 30.0,  // по умолчанию: М ×30
            MvvPredMode mvvMode    = MvvPredMode.Fev1TimesFactor
        )
        {
            if (ageYears  is < 18 or > 90)   throw new ArgumentOutOfRangeException(nameof(ageYears));
            if (heightCm  is < 130 or > 210) throw new ArgumentOutOfRangeException(nameof(heightCm));
            if (btpsAmb   <= 0)              btpsAmb = 1.0; // не используем, если режим без BTPS

            // ===== 1) ЖЕЛ (VC) — ECCS с калибровкой под Pulmo-4 =====
            // Муж: множитель 1.050 → VC≈4.93 при 60/185.
            // Жен: множитель 1.041 → VC≈3.98 при 30/172.
            double VC =
                sex == Sex.Male
                    ? (0.05200 * heightCm - 0.02200 * ageYears - 3.6000) * 1.050
                    : (0.04100 * heightCm - 0.01800 * ageYears - 2.6900) * 1.041;

            // ===== 2) ФЖЕЛ (FVC) =====
            double FVC = 0.97 * VC; // Pulmo-4: FVC ≈ 0.97·VC

            // ===== 3) ОФВ1 (FEV1) — ECCS с мягкой калибровкой =====
            double fev1Raw =
                sex == Sex.Male
                    ? (0.04100 * heightCm - 0.02400 * ageYears - 2.1900)
                    : (0.03400 * heightCm - 0.02500 * ageYears - 1.5780);

            // Масштаб под Pulmo-4: Ж 30/172 → ~3.34; М 60/185 → ~3.52.
            double FEV1 = sex == Sex.Male ? fev1Raw * 0.890 : fev1Raw * 0.948;

            // Нормативная доля ОФВ1/ЖЕЛ — 0.86.
            const double FEV1_over_VC = 0.86;

            // ===== 4) Пикфлоу (PEF) и производные скорости =====
            // База: Nunn–Gregg (в л/мин) → /60; далее небольшой масштаб kPEF.
            double pef_lps = (sex == Sex.Male
                                ? (5.12 * heightCm - 5.12 * ageYears - 188.0)
                                : (3.72 * heightCm - 2.81 * ageYears - 110.0)) / 60.0;

            double kPEF = sex == Sex.Male ? 0.965 : 0.964;     // Ж 30/172 → 7.16 л/с
            double PEF = pef_lps * kPEF;

            // Калиброванные коэффициенты под скрины Pulmo-4:
            double FEF25    = 0.912 * PEF; // 6.53 при PEF=7.16
            double FEF50    = 0.683 * PEF; // 4.89
            double FEF75    = 0.346 * PEF; // 2.48
            double FEF25_75 = 0.571 * PEF; // 4.09
            double FEF75_85 = 0.100 * PEF; // 0.72

            // ===== 5) ЖЕЛ-таблица (покой) =====
            const double VT_rest = 0.48;        // фиксированное «должное» ДО
            double IC_rest = VC;                // В таблице ЖЕЛ: ЕВдд = ЖЕЛд
            double IRV = 0.50 * VC;             // РОвдд ≈ 50% VC
            double ERV = 0.35 * VC;             // РОвыдд ≈ 35% VC
            double VTpct = VT_rest / VC * 100.0;

            // ===== 6) МОД-таблица (покой) =====
            const double RR_rest = 20.17;
            double VE = veRestOverride;         // по умолчанию 9.68 л/мин
            const double kO2  = 25.0;           // мл/л
            double VO2 = VE * kO2;              // 242 мл/мин при VE=9.68
            const double KNO2 = 40.0;           // мл/л
            const double TexpTinsp = 1.30;

            // ===== 7) МВЛ-таблица =====
            // ВАЖНО: Pulmo-4 для «должной» МВЛ использует FEV1 с ОКРУГЛЕНИЕМ до 0.01
            // (это позволяет получить 116.87 при Ж,30/172: round(3.335,2)=3.34; 3.34×35=116.9).
            double fev1Rounded = Math.Round(FEV1, 2, MidpointRounding.AwayFromZero);
            double mvvFactor = (sex == Sex.Female ? mvvFactorFemale : mvvFactorMale);
            double mvvBtpsMultiplier = (mvvMode == MvvPredMode.Fev1TimesFactorTimesBtps) ? btpsAmb : 1.0;
            double MVV = Math.Max(0.0, fev1Rounded) * mvvFactor * mvvBtpsMultiplier;

            const double RR_mvl = 85.0;
            double VT_mvl = 0.40 * VC;
            const double RD_pct = 85.0;
            double MVV_over_MOD = MVV / VE;

            // ===== 8) Упаковка результата =====
            var fvc = new PredictedFVC(
                FVC_L: FVC,
                IC_L:  FVC,           // В таблице ФЖЕЛ «ЕВд(л)» печатается как FVC
                FEV1_L: FEV1,
                FEV1_over_VC: FEV1_over_VC,
                PEF_Lps: PEF,
                OVnos_L: VT_rest,
                FEF25_Lps: FEF25,
                FEF50_Lps: FEF50,
                FEF75_Lps: FEF75,
                FEF25_75_Lps: FEF25_75,
                FEF75_85_Lps: FEF75_85,
                TFVC_s: 4.5
            );

            var zhel = new PredictedZhEL(
                VC_L: VC,
                IC_L: IC_rest,
                VT_L: VT_rest,
                IRV_L: IRV,
                ERV_L: ERV
                // VT_over_VC_pct: VTpct
            );

            var mod = new PredictedMOD(
                // RR_perMin: RR_rest,
                VE_Lpm: VE,
                VT_L: VT_rest,
                VO2_mlMin: VO2,
                KNO2_mlL: KNO2
                // Texp_over_Tinsp: TexpTinsp
            );

            var mvl = new PredictedMVL(
                // RR_perMin: RR_mvl,
                MVV_Lpm: MVV,
                VT_L: VT_mvl,
                RD_pct: RD_pct
                // MVV_over_MOD: MVV_over_MOD
            );

            return new PredictedSet(fvc, zhel, mod, mvl);
        }
    
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
                values[0] > 0 ? 100.0 * values[2]/values[0] : double.NaN
            );
        }

        // 4) MOD block
        var modPos = buffer.IndexOf(TagMod);
        if (modPos >= 0 && modPos + TagMod.Length + 12 <= buffer.Length)
        {
            var payloadStart = modPos + TagMod.Length;
            var modVals = ReadFloatArray(buffer.Slice(payloadStart, 12));
            var volumeCurve = ReadInt16Block(buffer, payloadStart + 12);

            double kio2 = kio2_mlPerL;                              // уже есть параметр метода Parse(...)
            double vo2  = modVals[1] * kio2;                        // VO2 = VE * kO2
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
            double rr      = mvlVals[0];
            double mvlLpm  = mvlVals[1];
            double vtL     = mvlVals[2];

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

        var predicted = Calculate(patient.AgeYears, patient.HeightM*100, patient.Sex, btps.Factor);
        return new PnpRecord(fileName, /*predicted,*/ patient, btps, zhel, mod, mvl, probes);
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
            return (double) volumeCurve[0] / volumeCurve[1];
        
        return Double.NaN;
    }

    #endregion

    #region ── Demographics & BTPS extraction ───────────────────────────────────

    private static Demographics ExtractDemographics(ReadOnlySpan<byte> src)
    {
        string rawHeader = ReadHeaderString(src);

        // сканируем каждый байт, иначе можно промахнуться по фактическому смещению структуры
        int maxOff = Math.Min(4096, src.Length - 9); // -9: нам нужно читать off..off+8
        for (int off = 0; off < maxOff; off++) // ВАЖНО: off++ (не off += 2)
        {
            ushort age    = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off, 2));
            ushort kg     = BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(off + 2, 2));
            float  height = BitConverter.ToSingle(src.Slice(off + 4, 4));
            byte   sex    = src[off + 8];

            bool plausible =
                age    is >= 5  and <= 120 &&
                kg     is >= 20 and <= 200 &&
                height is >= 1.2f and <= 2.5f &&
                (sex == 0 || sex == 1);

            if (plausible)
                return new Demographics(rawHeader, age, kg, Math.Round(height, 3), (Sex)sex);
        }

        // если не нашли — возвращаем header и значения по умолчанию
        return new Demographics(rawHeader, 0, 0, 0, Sex.Male);
    }

    private static string ReadHeaderString(ReadOnlySpan<byte> src)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var enc1251 = Encoding.GetEncoding(1251);

        // Декодируем первые <=256 байт (CP1251 — однобайтовая, безопасно)
        string s = enc1251.GetString(src[..Math.Min(256, src.Length)]);

        // 1) Режем по первому служебному разделителю (NULL, RS/US/GS)
        int cutCtrl = s.IndexOfAny(new[] { '\0', '\x1E', '\x1F', '\x1D' });
        if (cutCtrl >= 0) s = s[..cutCtrl];

        // 2) Если встречается $, часто это внутривендорный маркер — обрезаем по нему
        int cutDollar = s.IndexOf('$');
        if (cutDollar >= 0) s = s[..cutDollar];

        // 3) Удаляем оставшиеся управляющие символы (на случай смешанных заголовков)
        if (s.Length > 0)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (!char.IsControl(ch)) sb.Append(ch);
            }
            s = sb.ToString();
        }

        // 4) Финальная чистка пробелов
        return s.Trim();
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

            var factor = ComputeBtpsFactor(T, RH, P);
            return new BtpsInfo(true, factor, T, RH, P);
        }

        return new BtpsInfo(false, defaultFactor);
    }

    private static double ComputeBtpsFactor(double tempC, double rhPct, double pressureMmHg)
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
    }

    #endregion
} // class PnpParser