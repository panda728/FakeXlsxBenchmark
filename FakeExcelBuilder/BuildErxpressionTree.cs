using System.Buffers;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace FakeExcelBuilder
{
    public class BuilderExpressionTree
    {
        //private const int XF_NORMAL = 0;
        private const int XF_WRAP_TEXT = 1;
        private const int XF_DATE = 2;
        private const int XF_DATETIME = 3;

        string _frozenTitleRow = @"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>";

        public void Run<T>(string fileName, IEnumerable<T> rows, bool writeTitle = true, bool columnAutoFit = true)
        {
            var workPath = Path.Combine("work", Guid.NewGuid().ToString());
            var workRelPath = Path.Combine(workPath, "_rels");
            if (!Directory.Exists(workPath))
                Directory.CreateDirectory(workPath);

            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            if (File.Exists(fileName))
                File.Delete(fileName);

            var tempPath = "Templete";
            var tempRelPath = Path.Combine(tempPath, "_rels");

            try
            {
#if DEBUG
                File.Copy(Path.Combine(tempPath, "[Content_Types].xml"), Path.Combine(workPath, "[Content_Types].xml"));
                File.Copy(Path.Combine(tempPath, "book.xml"), Path.Combine(workPath, "book.xml"));
                File.Copy(Path.Combine(tempPath, "styles.xml"), Path.Combine(workPath, "styles.xml"));
                File.Copy(Path.Combine(tempRelPath, ".rels"), Path.Combine(workRelPath, ".rels"));
                File.Copy(Path.Combine(tempRelPath, "book.xml.rels"), Path.Combine(workRelPath, "book.xml.rels"));
                using (var fs = CreateStream(Path.Combine(workPath, "sheet.xml")))
                    CreateSheet(fs, rows, writeTitle, columnAutoFit);
                using (var fs = CreateStream(Path.Combine(workPath, "strings.xml")))
                    CreateStrings(fs);
                ZipFile.CreateFromDirectory(workPath, fileName);
#else
                using (var fs = CreateStream(Path.Combine(workPath, "sheet.xml")))
                    CreateSheet(fs, rows, writeTitle, columnAutoFit);
                using (var fs = CreateStream(Path.Combine(workPath, "strings.xml")))
                    CreateStrings(fs);
#endif
            }
            catch
            {
                throw;
            }
            finally
            {
#if DEBUG
#else
                try
                {
                    Directory.Delete(workPath, true);
                }
                catch { }
#endif
            }
        }

        private static Stream CreateStream(string fileName)
        {
#if DEBUG
            return new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);
#else
            return new FakeStream();
#endif
        }

        readonly byte[] _newLine = Encoding.UTF8.GetBytes(Environment.NewLine);
        readonly byte[] _rowTag1 = Encoding.UTF8.GetBytes("<row>");
        readonly byte[] _rowTag2 = Encoding.UTF8.GetBytes("</row>");

        internal class PropCache
        {
            public PropCache(Type type, PropertyInfo p)
            {
                Name = p.Name;

                var target = Expression.Parameter(typeof(object), p.Name);
                var lambda = Expression.Lambda<Func<object, object>>(
                    Expression.Convert(
                        Expression.PropertyOrField(
                            Expression.Convert(
                                target
                                , type)
                            , p.Name)
                        , typeof(object))
                    , target);

                Getter = lambda.Compile();
            }

            public string Name { get; init; }
            public Func<object, object> Getter { get; init; }
            public int Length { get; set; } = 0;
        }

        private void CreateSheet<T>(Stream stream, IEnumerable<T> rows, bool writeTitle, bool columnAutoFit)
        {
            SharedStringsClear();

            var properties = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(x => new PropCache(typeof(T), x))
                .ToArray()
                .AsSpan();

            @"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">"
                .WriteToStream(stream);
            _newLine.WriteToStream(stream);

            if (writeTitle)
            {
                _frozenTitleRow.WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }

            //foreach(var r in rows.Take(10))
            //{
            //    foreach (var p in properties)
            //        Console.WriteLine(p.Getter(r));
            //}

            if (columnAutoFit)
            {
                foreach (var p in properties)
                {
                    var maxLength = rows
                        .Take(100)
                        .Select(r =>
                        {
                            if (r == null) return "";
                            return p.Getter(r)?.ToString() ?? "";
                        })
                        .Max(s => s.Length);
                    p.Length = Math.Min(
                        Math.Max(maxLength, p.Name.Length) + 2,
                        100);
                }

                var i = 0;
                "<cols>".WriteToStream(stream);
                foreach (var p in properties)
                {
                    ++i;
                    @$"<col min=""{i}"" max =""{i}"" width =""{p.Length:0.0}"" bestFit =""1"" customWidth =""1"" />"
                        .WriteToStream(stream);
                }
                "</cols>".WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }

            "<sheetData>".WriteToStream(stream);
            _newLine.WriteToStream(stream);

            if (writeTitle)
            {
                _rowTag1.WriteToStream(stream);
                foreach (var p in properties)
                    GetColumnXml(p.Name).WriteToStream(stream);
                _rowTag2.WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }

            foreach (var r in rows)
            {
                if (r == null)
                    continue;

                _rowTag1.WriteToStream(stream);
                foreach (var p in properties)
                    GetColumnXml(p.Getter(r)).WriteToStream(stream);
                _rowTag2.WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }

            "</sheetData>".WriteToStream(stream);
            _newLine.WriteToStream(stream);
            "</worksheet>".WriteToStream(stream);
            _newLine.WriteToStream(stream);
        }

#region c Builder
        static readonly string _emptyColumn = "<c></c>";
        static int _index = 0;

        public static Dictionary<string, int> SharedStrings { get; } = new();
        public static void SharedStringsClear()
        {
            _index = 0;
            SharedStrings.Clear();
        }

        static string GetObjectColumnXml(object o)
            => @$"<c t=""s""><v>{GetSharedStringIndex(o?.ToString() ?? "")}</v></c>";

        static string GetStringColumnXml(string s)
        {
            if (string.IsNullOrEmpty(s))
                return _emptyColumn;

            var index = GetSharedStringIndex(s);
            return s.Contains(Environment.NewLine)
                ? @$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>{index}</v></c>"
                : @$"<c t=""s""><v>{index}</v></c>";
        }

        static int GetSharedStringIndex(string s)
        {
            if (SharedStrings.ContainsKey(s))
                return SharedStrings[s];

            SharedStrings.Add(s, _index);
            return _index++;
        }

        static string GetBooleanColumnXml(bool b) => @$"<c t=""b""><v>{b}</v></c>";
        static string GetNumberColumnXml<T>(T o) => @$"<c t=""n""><v>{o}</v></c>";
        static string GetDateColumnXml<T>(T o)
        {
            if (o is DateTime dateTime)
            {
                return dateTime.Hour == 0 && dateTime.Minute == 0 && dateTime.Second == 0
                    ? @$"<c t=""d"" s=""{XF_DATE}""><v>{dateTime:yyyy-MM-ddTHH:mm:ss}</v></c>"
                    : @$"<c t=""d"" s=""{XF_DATETIME}""><v>{dateTime:yyyy-MM-ddTHH:mm:ss}</v></c>";
            }
            else if (o is DateOnly dateOnly)
                return @$"<c t=""d"" s=""{XF_DATE}""><v>{dateOnly:yyyy-MM-ddTHH:mm:ss}</v></c>";
            else if (o is TimeOnly timeOnly)
                return @$"<c t=""d"" s=""{XF_DATETIME}""><v>{timeOnly:yyyy-MM-ddTHH:mm:ss}</v></c>";
            else
                return _emptyColumn;
        }

        public static string GetColumnXml(object? o)
        {
            if (o == null) return _emptyColumn;

            var t = o.GetType();
            if (t.IsGenericType || t.IsInterface || t.IsArray) return _emptyColumn;

            if (o is string s) return GetStringColumnXml(s);
            if (o is int i) return GetNumberColumnXml(i);
            if (o is long l) return GetNumberColumnXml(l);
            if (o is float f) return GetNumberColumnXml(f);
            if (o is double d) return GetNumberColumnXml(d);
            if (o is decimal de) return GetNumberColumnXml(de);
            if (o is bool b) return GetBooleanColumnXml(b);
            if (o is Guid g) return GetStringColumnXml($"{g}");
            if (o is DateTime dateTime) return GetDateColumnXml(dateTime);
            if (o is DateOnly dateOnly) return GetDateColumnXml(dateOnly);
            if (o is TimeOnly timeOnly) return GetDateColumnXml(timeOnly);

            return GetObjectColumnXml(o);
        }
#endregion

        private void CreateStrings(Stream stream)
        {
            $@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""{SharedStrings.Count}"">"
                .WriteToStream(stream);
            _newLine.WriteToStream(stream);

            var tagA = Encoding.UTF8.GetBytes("<si><t>").AsSpan();
            var tagB = Encoding.UTF8.GetBytes("</t></si>").AsSpan();
            foreach (var s in SharedStrings)
            {
                tagA.WriteToStream(stream);
                s.Key.WriteToStream(stream);
                tagB.WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }
            "</sst>".WriteToStream(stream);
            _newLine.WriteToStream(stream);
        }
    }
}
