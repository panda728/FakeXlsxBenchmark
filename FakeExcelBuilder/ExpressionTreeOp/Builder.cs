using System.Buffers;
using System.Collections.Concurrent;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace FakeExcelBuilder.ExpressionTreeOp
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

        readonly byte[] _newLine = Encoding.UTF8.GetBytes(Environment.NewLine);
        readonly byte[] _rowTag1 = Encoding.UTF8.GetBytes("<row>");
        readonly byte[] _rowTag2 = Encoding.UTF8.GetBytes("</row>");
        readonly byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

        private const int COLUMN_WIDTH_MAX = 100;
        private const int COLUMN_WIDTH_MARGIN = 2;

        static readonly ConcurrentDictionary<Type, FormatterHelper[]> _dic = new();
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
                using (var fsSheet = CreateStream(Path.Combine(workPath, "sheet.xml")))
                using (var fsString = CreateStream(Path.Combine(workPath, "strings.xml")))
                {
                    CreateSheet(rows, fsSheet, fsString, writeTitle, columnAutoFit);
                }
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

        public void Compile<T>() => _ = ParseFormatters<T>();
        private static FormatterHelper[] ParseFormatters<T>()
            => _dic.GetOrAdd(typeof(T),
                typeof(T)
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .Select(x => new FormatterHelper()
                    {
                        Name = x.Name,
                        Formatter = PropertyInfoExtensions.GenerateEncodedGetterLambda<T>(x)
                    })
                    .ToArray()
            );

        public void CreateSheet<T>(IEnumerable<T> rows, Stream fsSheet, Stream fsString, bool writeTitle, bool autoFitColumns)
        {
            using var writer = new ArrayPoolBufferWriter(1024);
            ColumnFormatter.SharedStringsClear();
            var formatters = ParseFormatters<T>().AsSpan();

            if (autoFitColumns)
            {
                foreach (var f in formatters)
                {
                    var maxLength = rows
                        .Take(100)
                        .Select(r => f.Formatter == null ? 0 : (int)f.Formatter(r, writer))
                        .Max(x => x);
                    f.MaxLength = Math.Min(
                        Math.Max(maxLength, f.Name.Length) + COLUMN_WIDTH_MARGIN,
                        COLUMN_WIDTH_MAX);
                }
                writer.Clear();
            }

            WriteHeader<T>(fsSheet, writeTitle, autoFitColumns);

            if (autoFitColumns)
            {
                var i = 0;
                fsSheet.Write(Encoding.UTF8.GetBytes("<cols>"));
                foreach (var f in formatters)
                {
                    ++i;
                    Encoding.UTF8.GetBytes(
                        @$"<col min=""{i}"" max =""{i}"" width =""{f.MaxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
                        writer);
                }
                Encoding.UTF8.GetBytes("</cols>", writer);
                writer.Write(_newLine);
            }
            Encoding.UTF8.GetBytes("<sheetData>", writer);
            writer.Write(_newLine);
            writer.CopyTo(fsSheet);

            if (writeTitle)
            {
                writer.Write(_rowTag1);
                foreach (var f in formatters)
                    ColumnFormatter.GetBytes(f.Name, writer);
                writer.Write(_rowTag2);
                writer.Write(_newLine);
                writer.CopyTo(fsSheet);
            }

            foreach (var row in rows)
            {
                if (row == null) continue;
                writer.Write(_rowTag1);
                foreach (var f in formatters)
                {
                    if (f.Formatter == null)
                        _ = ColumnFormatter.WriteEmptyCoulumn(writer);
                    else
                        _ = f.Formatter(row, writer);
                }
                writer.Write(_rowTag2);
                writer.Write(_newLine);
                writer.CopyTo(fsSheet);
            }

            fsSheet.Write(Encoding.UTF8.GetBytes("</sheetData>"));
            fsSheet.Write(_newLine);
            fsSheet.Write(Encoding.UTF8.GetBytes("</worksheet>"));
            fsSheet.Write(_newLine);

            WriteSharedStrings(fsString, ColumnFormatter.SharedStrings);
            ColumnFormatter.SharedStringsClear();
        }

        void WriteHeader<T>(Stream stream, bool writeTitle, bool autoFitColumns)
        {
            stream.Write(
                Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">")
            );
            stream.Write(_newLine);

            if (writeTitle)
            {
                stream.Write(_frozenTitleRow);
                stream.Write(_newLine);
            }
        }

        private void WriteSharedStrings(Stream stream, Dictionary<string, int> sharedStrings)
        {
            using var writer = new ArrayPoolBufferWriter(1024);
            Encoding.UTF8.GetBytes($@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""{sharedStrings.Count}"">"
                , writer);
            writer.Write(_newLine);

            var tagA = Encoding.UTF8.GetBytes("<si><t>").AsSpan();
            var tagB = Encoding.UTF8.GetBytes("</t></si>").AsSpan();
            foreach (var s in sharedStrings)
            {
                writer.Write(tagA);
                Encoding.UTF8.GetBytes(s.Key, writer);
                writer.Write(tagB);
                writer.Write(_newLine);
                writer.CopyTo(stream);
            }
            Encoding.UTF8.GetBytes("</sst>", writer);
            writer.Write(_newLine);
            writer.CopyTo(stream);
        }
    }
}
