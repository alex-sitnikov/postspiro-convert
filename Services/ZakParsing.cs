// ZakParser.cs — .NET 5/6+, C# 9
// Объект на файл: один Section, одна Area, Patient, Measurements[], Conclusion[], Extras.

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace SpiroUI.Services
{
    public record Measurement(
        string Key,
        string Side, // "L" | "R" | "—"
        double? Value,
        string Unit,
        string Raw
    );

    public record PatientData(
        string? FullName,
        int? Age,
        string? Sex, // "Ж" | "М"
        int? Height,
        int? Weight,
        DateTime? Date,
        string? Comment
    );

    public record ConclusionItem(
        string Key,
        string Value,
        string? Note,
        double? DeltaPercent,
        string? DeltaDirection // "больше" | "меньше" | null
    );

    public record Extras(
        double? AsymmetryCoeffPercent,
        string? AsymmetryNormText,
        string? AsymmetryQualitative,
        string? AsymmetryDominanceCode, // "S>D" / "D>S"
        string? AsymmetryDominanceSide, // "Left" / "Right"
        int? HeartRateBpm,
        int? HeartRateLow,
        int? HeartRateHigh
    );

    public record ZakRecord(
        string FileName,
        string Section, // например: "РЕОЭНЦЕФАЛОГРАФИЯ"
        string? Area, // например: "Фронто - мастоидальная область (FM)"
        PatientData Patient,
        List<Measurement> Measurements,
        List<ConclusionItem> Conclusion,
        Extras Extra
    );

    public static class ZakParser
    {
        private const char BrokenBar = '\u00A6'; // ¦

        static ZakParser()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // ---------- API ----------
        public static ZakRecord ParseFile(string path)
        {
            var text = ReadText(path);
            var lines = Normalize(text).Split('\n');

            var (section, area) = GetSectionArea(lines);
            var patient = ParsePatient(lines);
            var measurements = ParseMeasurements(lines);
            var conclusion = ParseConclusion(text, section, area);
            var extra = ParseExtras(text);

            return new ZakRecord(Path.GetFileName(path), section, area, patient, measurements, conclusion, extra);
        }

        public static string ToMetricsLongCsv(IEnumerable<ZakRecord> records)
        {
            var cols = new[] { "File", "Section", "Area", "Key", "Side", "Value", "Unit", "Raw" };
            var sb = new StringBuilder().AppendLine(string.Join(",", cols));
            foreach (var r in records)
            {
                // Measurements
                foreach (var m in r.Measurements)
                {
                    var row = new[]
                    {
                        Csv(r.FileName), Csv(r.Section), Csv(r.Area ?? ""), Csv(m.Key), Csv(m.Side),
                        m.Value?.ToString(CultureInfo.InvariantCulture) ?? "", Csv(m.Unit), Csv(m.Raw)
                    };
                    sb.AppendLine(string.Join(",", row));
                }

                // Patient
                foreach (var (k, v) in PatientKvp(r.Patient))
                {
                    var row = new[] { Csv(r.FileName), Csv("Patient"), "", Csv(k), Csv("—"), Csv(v), "", Csv(v) };
                    sb.AppendLine(string.Join(",", row));
                }

                // Extras as rows (для Excel-наблюдения)
                foreach (var (k, v, unit, raw) in ExtrasRows(r.Extra))
                {
                    var row = new[]
                    {
                        Csv(r.FileName), Csv(r.Section), Csv(r.Area ?? ""), Csv(k), Csv("—"), Csv(v), Csv(unit ?? ""),
                        Csv(raw ?? "")
                    };
                    sb.AppendLine(string.Join(",", row));
                }
            }

            return sb.ToString();
        }

        public static string ToConclusionCsv(IEnumerable<ZakRecord> records)
        {
            var cols = new[] { "File", "Section", "Area", "Key", "Value", "Note", "DeltaPercent", "DeltaDirection" };
            var sb = new StringBuilder().AppendLine(string.Join(",", cols));
            foreach (var r in records)
            foreach (var c in r.Conclusion)
            {
                var row = new[]
                {
                    Csv(r.FileName), Csv(r.Section), Csv(r.Area ?? ""), Csv(c.Key), Csv(c.Value),
                    Csv(c.Note ?? ""), c.DeltaPercent?.ToString(CultureInfo.InvariantCulture) ?? "",
                    Csv(c.DeltaDirection ?? "")
                };
                sb.AppendLine(string.Join(",", row));
            }

            return sb.ToString();
        }

        // ---------- Internals ----------

        // Строгое распознавание строки-шапки двух колонок, например:
        // "Левая голень : Правая голень" или "Левая сторона : Правая сторона".
        // Требуем, чтобы строка была "чистой" (без процентов, скобок, цифр и т.п.).
        // Строго наличие обеих сторон в одной строке вида "Левая ... : Правая ..."
        // Разрешаем любые символы, кроме двоеточия, внутри названий колонок.
        private static bool TryGetTwoColumnHeader(string line, out string left, out string right)
        {
            left = right = string.Empty;
            if (string.IsNullOrWhiteSpace(line)) return false;

            var m = Regex.Match(
                line,
                @"^\s*[_\-\s]*\s*(Лев[^:]{1,80}?)\s*:\s*(Прав[^:]{1,80}?)\s*$",
                RegexOptions.IgnoreCase
            );
            if (!m.Success) return false;

            left  = m.Groups[1].Value.Trim(' ', '.', '\t');
            right = m.Groups[2].Value.Trim(' ', '.', '\t');
            return
                Regex.IsMatch(left,  @"^Лев",  RegexOptions.IgnoreCase) &&
                Regex.IsMatch(right, @"^Прав", RegexOptions.IgnoreCase);
        }
        
        private static (string left, string right) SplitNoteLR(string s)
        {
            s = s.Trim();
            var m = Regex.Match(s, @"\)\s*:\s*\(?");
            if (m.Success)
            {
                var left = s.Substring(0, m.Index);
                var right = s.Substring(m.Index + m.Length);
                return (left.Trim(), right.Trim());
            }

            var idx = s.IndexOf(':');
            if (idx >= 0)
                return (s.Substring(0, idx).Trim(), s[(idx + 1)..].Trim());
            return (s, "");
        }

        private static string ReadText(string path)
        {
            var bytes = File.ReadAllBytes(path);
            try
            {
                return Encoding.GetEncoding(1251).GetString(bytes);
            }
            catch
            {
                return Encoding.UTF8.GetString(bytes);
            }
        }

        private static string Normalize(string t)
        {
            t = t.Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = t.Split('\n').Select(x => x.TrimEnd());
            return string.Join("\n", lines);
        }

        private static string CollapseSpacedCaps(string s)
        {
            var parts = Regex.Split(s, "(\\s+)");
            var sb = new StringBuilder();
            bool IsCap(string t) => t.Length == 1 && IsUpperLetter(t[0]);
            int i = 0;
            while (i < parts.Length)
            {
                var tok = parts[i];
                if (i + 2 < parts.Length && IsCap(tok) && parts[i + 1] == " " && IsCap(parts[i + 2]))
                {
                    var letters = new StringBuilder(tok);
                    i += 2;
                    while (i < parts.Length && parts[i - 1] == " " && IsCap(parts[i]))
                    {
                        letters.Append(parts[i]);
                        i += 2;
                    }

                    sb.Append(letters.ToString());
                    if (i - 1 < parts.Length && Regex.IsMatch(parts[i - 1] ?? "", @"\s{2,}")) sb.Append(' ');
                }
                else
                {
                    sb.Append(tok);
                    i++;
                }
            }

            return Regex.Replace(sb.ToString(), " {2,}", " ").Trim();
        }

        private static bool IsUpperLetter(char ch) =>
            (ch >= 'A' && ch <= 'Z') || (ch >= 'А' && ch <= 'Я') || ch == 'Ё';

        private static string Csv(string s)
        {
            if (s == null) return "";
            if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
                return '"' + s.Replace("\"", "\"\"") + '"';
            return s;
        }

        private static IEnumerable<(string k, string v)> PatientKvp(PatientData p)
        {
            var dict = new Dictionary<string, string?>()
            {
                ["FullName"] = p.FullName, ["Age"] = p.Age?.ToString(),
                ["Sex"] = p.Sex, ["Height"] = p.Height?.ToString(), ["Weight"] = p.Weight?.ToString(),
                ["Date"] = p.Date?.ToString("yyyy-MM-dd"), ["Comment"] = p.Comment
            };
            foreach (var kv in dict)
                if (!string.IsNullOrEmpty(kv.Value))
                    yield return (kv.Key, kv.Value!);
        }

        private static IEnumerable<(string k, string v, string? unit, string? raw)> ExtrasRows(Extras e)
        {
            if (e.AsymmetryCoeffPercent is not null)
                yield return ("Коэффициент асимметрии",
                    e.AsymmetryCoeffPercent.Value.ToString(CultureInfo.InvariantCulture), "%", null);
            if (!string.IsNullOrWhiteSpace(e.AsymmetryQualitative))
                yield return ("Асимметрия кровенаполнения (кач.)", e.AsymmetryQualitative!, null, null);
            if (!string.IsNullOrWhiteSpace(e.AsymmetryDominanceCode))
                yield return ("Доминирование (S/D)", e.AsymmetryDominanceCode!, null, null);
            if (e.HeartRateBpm is not null)
                yield return ("ЧСС", e.HeartRateBpm.Value.ToString(CultureInfo.InvariantCulture), "в мин.", null);
            if (e.HeartRateLow is not null && e.HeartRateHigh is not null)
                yield return ("ЧСС диапазон", $"{e.HeartRateLow}-{e.HeartRateHigh}", null, null);
            if (!string.IsNullOrWhiteSpace(e.AsymmetryNormText))
                yield return ("Норма (для коэф. асим.)", e.AsymmetryNormText!, null, null);
        }

        private static PatientData ParsePatient(string[] lines)
        {
            string? name = null, sex = null, comment = null;
            int? age = null, height = null, weight = null;
            DateTime? date = null;

            foreach (var ln in lines)
            {
                if (ln.Contains("Фамилия,имя,отчество"))
                {
                    var idx = ln.IndexOf(':');
                    if (idx >= 0) name = ln[(idx + 1)..].Trim();
                }

                if (Regex.IsMatch(ln, "Возраст\\s*:"))
                {
                    var mAge = Regex.Match(ln, "Возраст\\s*:\\s*([0-9]*)");
                    if (mAge.Success && int.TryParse(mAge.Groups[1].Value, out var a)) age = a;
                    var mSex = Regex.Match(ln, "Пол\\s*:\\s*([ЖМ])");
                    if (mSex.Success) sex = mSex.Groups[1].Value;
                    var mH = Regex.Match(ln, "Рост\\s*:\\s*([0-9]*)");
                    if (mH.Success && int.TryParse(mH.Groups[1].Value, out var h)) height = h;
                    var mW = Regex.Match(ln, "Вес\\s*:\\s*([0-9]*)");
                    if (mW.Success && int.TryParse(mW.Groups[1].Value, out var w)) weight = w;
                }

                if (Regex.IsMatch(ln, "Дата\\s*:"))
                {
                    var m = Regex.Match(ln, "Дата\\s*:\\s*([0-9]{1,2})\\s+([0-9]{1,2})\\s+([0-9]{2,4})(.*)$");
                    if (m.Success)
                    {
                        var dd = int.Parse(m.Groups[1].Value);
                        var mm = int.Parse(m.Groups[2].Value);
                        var yyStr = m.Groups[3].Value.PadLeft(4, '0');
                        if (yyStr.Length == 2) yyStr = "20" + yyStr;
                        if (int.TryParse(yyStr, out var y))
                        {
                            try
                            {
                                date = new DateTime(y, mm, dd);
                            }
                            catch
                            {
                            }
                        }

                        comment = Regex.Replace(m.Groups[4].Value, "^\\.*\\s*", "").Trim();
                        if (string.IsNullOrWhiteSpace(comment)) comment = null;
                    }
                }
            }

            return new PatientData(name, age, sex, height, weight, date, comment);
        }

        private static (string section, string? area) GetSectionArea(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                var s = lines[i].Trim();
                var s2 = Regex.Replace(s, "\\s+", "");
                bool looks = s2.Length >= 6 && (s2.Contains("РЕО") || s2.Contains("РВГ") ||
                                                s2.Contains("РЕОВАЗОГРАФИЯ") || s2.Contains("РЕОЭНЦЕФАЛОГРАФИЯ"));
                if (!looks) continue;

                string section = CollapseSpacedCaps(s);
                string? area = null;
                if (i + 1 < lines.Length)
                {
                    var nxt = lines[i + 1].Trim();
                    if (LikelyAreaLine(nxt)) area = nxt;
                }

                return (section, area);
            }

            return ("Unknown", null);
        }

        private static bool LikelyAreaLine(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (s.IndexOf("область", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (s.Contains("(") && s.Contains(")") && s.Length < 100) return true;
            if (s.Contains(BrokenBar)) return false;
            var decor = "*=_-¦";
            bool allDecor = s.Trim().All(ch => decor.Contains(ch));
            return s.Length < 80 && !allDecor;
        }

        private static List<Measurement> ParseMeasurements(string[] lines)
        {
            var list = new List<Measurement>();
            foreach (var ln in lines)
            {
                if (ln.IndexOf(BrokenBar) >= 0 && ln.Count(c => c == BrokenBar) >= 3)
                {
                    var parts = ln.Split(BrokenBar).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
                    if (parts.Count < 3) continue;
                    var label = CollapseSpacedCaps(parts[0]);
                    if (Regex.IsMatch(label, "Основные", RegexOptions.IgnoreCase)) continue;
                    var L = parts[^2];
                    var R = parts[^1];

                    if (Regex.IsMatch(L, "\\d"))
                    {
                        var (v, u, raw) = ParseValueUnit(L);
                        list.Add(new Measurement(label, "L", v, u, raw));
                    }

                    if (Regex.IsMatch(R, "\\d"))
                    {
                        var (v, u, raw) = ParseValueUnit(R);
                        list.Add(new Measurement(label, "R", v, u, raw));
                    }
                }
            }

            return list;
        }

        private static (double? v, string unit, string raw) ParseValueUnit(string cell)
        {
            var raw = cell.Trim();
            var m = Regex.Match(raw.Trim(' ', '.'), @"([-+]?\d+(?:[.,]\d+)?)\s*([^\d\s+-].*)?$",
                RegexOptions.CultureInvariant);
            if (m.Success)
            {
                var val = m.Groups[1].Value.Replace(',', '.');
                if (double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                    return (d, (m.Groups[2].Value ?? "").Trim(), raw);
                return (null, (m.Groups[2].Value ?? "").Trim(), raw);
            }

            return (null, "", raw);
        }

        private static List<ConclusionItem> ParseConclusion(string text, string section, string? area)
        {
            var result = new List<ConclusionItem>();
            var m = Regex.Match(text, @"(З\s*А\s*К\s*Л\s*Ю\s*Ч\s*Е\s*Н\s*И\s*Е|РЕЗЮМЕ)\s*:?", RegexOptions.IgnoreCase);
            if (!m.Success) return result;

            var block = text[(m.Index + m.Length)..];
            var lines = Normalize(block).Split('\n');

            // Если внутри заключения встретится "В области ...", перезапишем area
            string? effectiveArea = area;
            foreach (var ln in lines.Take(6))
                if (Regex.IsMatch(ln, "В\\s+области", RegexOptions.IgnoreCase))
                {
                    var a = Regex.Split(ln, "В\\s+области", RegexOptions.IgnoreCase)[1].Trim(' ', ':');
                    effectiveArea = JoinSingleLetterGroups(a).Trim(' ', '.', ':', '-');
                    break;
                }

            // Пытаемся найти "чистую" шапку двух колонок (строка вида "Левая … : Правая …").
            string? colLeft = null, colRight = null;
            for (int j = 0; j < Math.Min(12, lines.Length); j++)
            {
                if (TryGetTwoColumnHeader(lines[j], out var l, out var r))
                {
                    colLeft = l; colRight = r;
                    break;
                }
            }

            string? currentBase = null;

            (string? note, double? dp, string? dd) ParseNote(string s)
            {
                var note = Regex.Replace(Regex.Replace(s.Trim(), @"^[\s:]*\(?\s*", ""), @"\)?\s*$", "");
                var mm = Regex.Match(note, @"([+-]?\d+(?:[.,]\d+)?)\s*%.*\b(больше|меньше)\s+нормы",
                    RegexOptions.IgnoreCase);
                double? dp = null;
                string? dd = null;
                if (mm.Success)
                {
                    if (double.TryParse(mm.Groups[1].Value.Replace(',', '.'), NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var d))
                        dp = d;
                    dd = mm.Groups[2].Value.ToLower();
                }

                return (string.IsNullOrWhiteSpace(note) ? null : note, dp, dd);
            }

            for (int i = 0; i < lines.Length; i++)
            {
                var ln = lines[i];
                if (Regex.IsMatch(ln, @"^\s*РЕЗЮМЕ\s*:", RegexOptions.IgnoreCase)) break;
                if (Regex.IsMatch(ln, @"^\s*-{5,}\s*$")) continue;

                // 1) Нумерованный пункт: "1. Кровенаполнение : <L> : <R>" ИЛИ "1. Кровенаполнение : <val>"
                var mnum = Regex.Match(ln, @"^\s*(\d+)\.\s*(.*)$");
                if (mnum.Success)
                {
                    currentBase = mnum.Groups[2].Value.Split(':')[0].Trim(' ', '.');

                    var parts = ln.Split(':');
                    var valL = parts.Length >= 2 ? parts[1].Trim() : "";
                    var valR = parts.Length >= 3 ? string.Join(":", parts.Skip(2)).Trim() : "";

                    if (colLeft != null && colRight != null && (valL.Length > 0 || valR.Length > 0))
                    {
                        // Двухколоночный вариант: создаём две записи
                        if (valL.Length > 0 && !valL.StartsWith("("))
                        {
                            var keyL = $"{currentBase} ({colLeft})";
                            var itemL = new ConclusionItem(keyL, valL.Trim(' ', '.'), null, null, null);

                            if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            {
                                var (rawL, rawR) = SplitNoteLR(lines[i + 1]);
                                var (nL, dpL, ddL) = ParseNote(rawL);
                                itemL = itemL with { Note = nL, DeltaPercent = dpL, DeltaDirection = ddL };
                            }

                            result.Add(itemL);
                        }

                        if (valR.Length > 0 && !valR.StartsWith("("))
                        {
                            var keyR = $"{currentBase} ({colRight})";
                            var itemR = new ConclusionItem(keyR, valR.Trim(' ', '.'), null, null, null);

                            if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            {
                                var (rawL, rawR) = SplitNoteLR(lines[i + 1]);
                                var (nR, dpR, ddR) = ParseNote(rawR);
                                itemR = itemR with { Note = nR, DeltaPercent = dpR, DeltaDirection = ddR };
                            }

                            result.Add(itemR);
                        }

                        // если была строка примечаний — мы её уже съели
                        if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            i++;

                        continue;
                    }
                    else
                    {
                        // Обычный (одноколоночный) вариант
                        var val = valL;
                        if (val.Length > 0 && !val.StartsWith("("))
                        {
                            var item = new ConclusionItem(currentBase, val.Trim(' ', '.'), null, null, null);
                            if (i + 1 < lines.Length && lines[i + 1].Contains("("))
                            {
                                var (n, dp, dd) = ParseNote(lines[i + 1]);
                                item = item with { Note = n, DeltaPercent = dp, DeltaDirection = dd };
                                i++;
                            }

                            result.Add(item);
                        }

                        continue;
                    }
                }

                // 2) Подпункт: "<подзаголовок> : <L> : <R>" ИЛИ "<подзаголовок> : <val>"
                if (currentBase != null && ln.Contains(":"))
                {
                    var parts = ln.Split(':');
                    var leftLabel = parts[0].Trim(' ', '.');
                    var valL = parts.Length >= 2 ? parts[1].Trim() : "";
                    var valR = parts.Length >= 3 ? string.Join(":", parts.Skip(2)).Trim() : "";
                    var label = string.IsNullOrWhiteSpace(leftLabel) || Regex.IsMatch(leftLabel, @"^[_\-]+$")
                        ? currentBase
                        : (currentBase + " " + leftLabel).Trim();

                    if (colLeft != null && colRight != null && (valL.Length > 0 || valR.Length > 0))
                    {
                        if (valL.Length > 0 && !valL.StartsWith("("))
                        {
                            var keyL = $"{label} ({colLeft})";
                            var itemL = new ConclusionItem(keyL, valL.Trim(' ', '.'), null, null, null);

                            if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            {
                                var (rawL, rawR) = SplitNoteLR(lines[i + 1]);
                                var (nL, dpL, ddL) = ParseNote(rawL);
                                itemL = itemL with { Note = nL, DeltaPercent = dpL, DeltaDirection = ddL };
                            }

                            result.Add(itemL);
                        }

                        if (valR.Length > 0 && !valR.StartsWith("("))
                        {
                            var keyR = $"{label} ({colRight})";
                            var itemR = new ConclusionItem(keyR, valR.Trim(' ', '.'), null, null, null);

                            if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            {
                                var (rawL, rawR) = SplitNoteLR(lines[i + 1]);
                                var (nR, dpR, ddR) = ParseNote(rawR);
                                itemR = itemR with { Note = nR, DeltaPercent = dpR, DeltaDirection = ddR };
                            }

                            result.Add(itemR);
                        }

                        if (i + 1 < lines.Length && (lines[i + 1].Contains("(") || lines[i + 1].Contains(":")))
                            i++;

                        continue;
                    }
                    else
                    {
                        var val = valL;
                        if (val.Length > 0 && !val.StartsWith("("))
                        {
                            var item = new ConclusionItem(label, val.Trim(' ', '.'), null, null, null);
                            if (i + 1 < lines.Length && lines[i + 1].Contains("("))
                            {
                                var (n, dp, dd) = ParseNote(lines[i + 1]);
                                item = item with { Note = n, DeltaPercent = dp, DeltaDirection = dd };
                                i++;
                            }

                            result.Add(item);
                        }
                    }
                }
            }

            return result;
        }

        private static Extras ParseExtras(string text)
        {
            double? coef = null;
            string? normText = null;
            var mCoef = Regex.Match(text,
                @"Коэфф\w*\s+асимметр\w*\s*[:=]\s*([+-]?\d+(?:[.,]\d+)?)\s*%?(?:\s*\(([^)]*Норма[^)]*)\))?",
                RegexOptions.IgnoreCase);
            if (mCoef.Success)
            {
                if (double.TryParse(mCoef.Groups[1].Value.Replace(',', '.'), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var d)) coef = d;
                normText = string.IsNullOrWhiteSpace(mCoef.Groups[2].Value) ? null : mCoef.Groups[2].Value.Trim();
            }

            string? qual = null, domCode = null, domSide = null;
            // Качественная асимметрия кровенаполнения + доминирование S>D/D>S (учитываем "в ПРЕДЕЛАХ НОРМЫ")
            var mQual = Regex.Match(
                text,
                @"(Асимметр[^\n()]*кровенаполнени[^\n()]*)\s*\(\s*([SDСД])\s*>\s*([SDСД])\s*\)",
                RegexOptions.IgnoreCase
            );
            if (mQual.Success)
            {
                qual = Regex.Replace(mQual.Groups[1].Value, @"\s{2,}", " ").Trim();

                string MapSD(string s)
                {
                    if (string.IsNullOrEmpty(s)) return s;
                    var ch = char.ToUpperInvariant(s[0]);
                    if (ch == 'S' || ch == 'С') return "S";
                    if (ch == 'D' || ch == 'Д') return "D";
                    return ch.ToString();
                }

                var leftC  = MapSD(mQual.Groups[2].Value);
                var rightC = MapSD(mQual.Groups[3].Value);
                domCode = $"{leftC}>{rightC}";
                domSide = leftC == "S" ? "Left" : leftC == "D" ? "Right" : null;
            }

            int? bpm = null, lo = null, hi = null;
            var mHr = Regex.Match(text,
                @"Частота\s+сердечных\s+сокращени[йя]\s*[:=]\s*(\d+)\s*(?:\((\d+)\s*[-–—]\s*(\d+)\))?\s*(?:в\s*мин\.?|уд/мин|bpm)?",
                RegexOptions.IgnoreCase);
            if (mHr.Success)
            {
                if (int.TryParse(mHr.Groups[1].Value, out var b)) bpm = b;
                if (int.TryParse(mHr.Groups[2].Value, out var l)) lo = l;
                if (int.TryParse(mHr.Groups[3].Value, out var h)) hi = h;
            }

            return new Extras(coef, normText, qual, domCode, domSide, bpm, lo, hi);
        }

        // Helper used in Conclusion area fallback
        private static string JoinSingleLetterGroups(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return s ?? "";
            var tokens = s.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            var res = new List<string>();
            for (int i = 0; i < tokens.Length;)
            {
                if (tokens[i].Length == 1 && IsUpperLetter(tokens[i][0]))
                {
                    var sb = new StringBuilder();
                    int j = i;
                    while (j < tokens.Length && tokens[j].Length == 1 && IsUpperLetter(tokens[j][0]))
                    {
                        sb.Append(tokens[j]);
                        j++;
                    }

                    res.Add(sb.ToString());
                    i = j;
                }
                else
                {
                    res.Add(tokens[i]);
                    i++;
                }
            }

            return string.Join(" ", res);
        }
    }
}