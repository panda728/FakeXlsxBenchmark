using System.Buffers;
using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace FakeExcelBuilder.Reflection
{

    public class Builder
    {
        byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
        byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

        byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets>
<sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/>
</sheets>
</workbook>");
        byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

        byte[] _styles = Encoding.UTF8.GetBytes(@"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<numFmts count=""2"">
<numFmt numFmtId=""1"" formatCode =""yyyy/mm/dd;@"" />
<numFmt numFmtId=""2"" formatCode =""yyyy/mm/dd\ hh:mm;@"" />
</numFmts>
<fonts count=""1"">
<font/>
</fonts>
<fills count=""1"">
<fill/>
</fills>
<borders count=""1"">
<border/>
</borders>
<cellStyleXfs count=""1"">
<xf/>
</cellStyleXfs>
<cellXfs count=""4"">
<xf/>
<xf><alignment wrapText=""true""/></xf>
<xf numFmtId=""1""  applyNumberFormat=""1""></xf>
<xf numFmtId=""2""  applyNumberFormat=""1""></xf>
</cellXfs>
</styleSheet>");

        //private const int XF_NORMAL = 0;
        private const int XF_WRAP_TEXT = 1;
        private const int XF_DATE = 2;
        private const int XF_DATETIME = 3;

        string _frozenTitleRow = @"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>";

        public async Task RunAsync<T>(string fileName, IEnumerable<T> rows, bool writeTitle = true, bool columnAutoFit = true)
        {
            var workPath = Path.Combine("work", Guid.NewGuid().ToString());
            var workRelPath = Path.Combine(workPath, "_rels");
            if (!Directory.Exists(workPath))
                Directory.CreateDirectory(workPath);

            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            if (File.Exists(fileName))
                File.Delete(fileName);

            try
            {
                using (var fs = CreateStream(Path.Combine(workPath, "[Content_Types].xml")))
                    await fs.WriteAsync(_contentTypes);
                using (var fs = CreateStream(Path.Combine(workRelPath, ".rels")))
                    await fs.WriteAsync(_rels);
                using (var fs = CreateStream(Path.Combine(workPath, "book.xml")))
                    await fs.WriteAsync(_book);
                using (var fs = CreateStream(Path.Combine(workRelPath, "book.xml.rels")))
                    await fs.WriteAsync(_bookRels);
                using (var fs = CreateStream(Path.Combine(workPath, "styles.xml")))
                    await fs.WriteAsync(_styles);
                using (var fs = CreateStream(Path.Combine(workPath, "sheet.xml")))
                    CreateSheet(fs, rows, writeTitle, columnAutoFit);
                using (var fs = CreateStream(Path.Combine(workPath, "strings.xml")))
                    CreateStrings(fs);
#if DEBUG
                ZipFile.CreateFromDirectory(workPath, fileName);
#else
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
            public PropCache(PropertyInfo p)
            {
                Name = p.Name;
                Accessor = p.GetAccessor();
            }

            public string Name { get; init; }
            public IAccessor Accessor { get; init; }
            public int Length { get; set; } = 0;
        }

        private void CreateSheet<T>(Stream stream, IEnumerable<T> rows, bool writeTitle, bool columnAutoFit)
        {
            SharedStringsClear();

            var properties = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(x => new PropCache(x))
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

            if (columnAutoFit)
            {
                foreach (var p in properties)
                {
                    var maxLength = rows
                        .Take(100)
                        .Select(r =>
                        {
                            if (r == null) return "";
                            return p.Accessor.GetValue(r)?.ToString() ?? "";
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
                    GetColumnXml(p.Accessor.GetValue(r)).WriteToStream(stream);
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
