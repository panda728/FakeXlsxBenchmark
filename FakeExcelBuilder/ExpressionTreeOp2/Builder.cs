using System.Buffers;
using System.Collections.Concurrent;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace FakeExcelBuilder.ExpressionTreeOp2
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
        readonly byte[] _rowStart = Encoding.UTF8.GetBytes("<row>");
        readonly byte[] _rowEnd = Encoding.UTF8.GetBytes("</row>");
        readonly byte[] _colStart = Encoding.UTF8.GetBytes("<cols>");
        readonly byte[] _colEnd = Encoding.UTF8.GetBytes("</cols>");
        readonly byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

        readonly byte[] _sheetStart = Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
        readonly byte[] _sheetEnd = Encoding.UTF8.GetBytes(@"</worksheet>");
        readonly byte[] _dataStart = Encoding.UTF8.GetBytes(@"<sheetData>");
        readonly byte[] _dataEnd = Encoding.UTF8.GetBytes(@"</sheetData>");

        private const int COLUMN_WIDTH_MAX = 100;
        private const int COLUMN_WIDTH_MARGIN = 2;

        public void Compile<T>() => _ = GetPropertiesCache<T>.Properties;

        private static class GetPropertiesCache<T>
        {
            static GetPropertiesCache()
            {
                Properties = typeof(T)
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .AsParallel()
                    .Select((p, i) => new FormatterHelper<T>(p, i))
                    .OrderBy(p => p.Index)
                    .ToArray();
            }
            public static readonly FormatterHelper<T>[] Properties;
        }

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
                    Formatter.SharedStringsClear();
                    CreateSheet(rows, fsSheet, writeTitle, columnAutoFit);
                    WriteSharedStrings(fsString, Formatter.SharedStrings);
                    Formatter.SharedStringsClear();
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

        void CalcCellStringLength<T>(IEnumerable<T> rows)
        {
            var formatters = GetPropertiesCache<T>.Properties.AsSpan();
            using var buffer = new ArrayPoolBufferWriter();
            foreach (var f in formatters)
            {
                var maxLength = rows
                    .Take(100)
                    .Select(r =>
                    {
                        if (r == null || f.Formatter == null) return 0;
                        var len = (int)f.Formatter(r, buffer);
                        buffer.Clear();
                        return len;
                    })
                    .Max(x => x);
                f.MaxLength = Math.Min(
                    Math.Max(maxLength, f.Name.Length) + COLUMN_WIDTH_MARGIN,
                    COLUMN_WIDTH_MAX);
            }
        }

        public void CreateSheet<T>(IEnumerable<T> rows, Stream fsSheet, bool writeTitle, bool autoFitColumns)
        {
            if (autoFitColumns)
                CalcCellStringLength(rows);

            fsSheet.Write(_sheetStart);
            fsSheet.Write(_newLine);

            if (writeTitle)
            {
                fsSheet.Write(_frozenTitleRow);
                fsSheet.Write(_newLine);
            }

            using var writer = new ArrayPoolBufferWriter();
            var formatters = GetPropertiesCache<T>.Properties.AsSpan();
            if (autoFitColumns)
            {
                var i = 0;
                writer.Write(_colStart);
                foreach (var f in formatters)
                {
                    ++i;
                    Encoding.UTF8.GetBytes(
                        @$"<col min=""{i}"" max =""{i}"" width =""{f.MaxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
                        writer);
                    writer.CopyTo(fsSheet);
                }
                writer.Write(_colEnd);
                writer.Write(_newLine);
                writer.CopyTo(fsSheet);
            }

            fsSheet.Write(_dataStart);
            fsSheet.Write(_newLine);

            if (writeTitle)
            {
                fsSheet.Write(_rowStart);
                foreach (var f in formatters)
                {
                    Formatter.Serialize(f.Name, writer);
                    writer.CopyTo(fsSheet);
                }
                fsSheet.Write(_rowEnd);
                fsSheet.Write(_newLine);
            }

            foreach (var row in rows)
            {
                if (row == null) continue;
                fsSheet.Write(_rowStart);
                foreach (var f in formatters)
                {
                    f.Formatter(row, writer);
                    writer.CopyTo(fsSheet);
                }

                fsSheet.Write(_rowEnd);
                fsSheet.Write(_newLine);
            }

            fsSheet.Write(_dataEnd);
            fsSheet.Write(_newLine);
            fsSheet.Write(_sheetEnd);
            fsSheet.Write(_newLine);
        }

        readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"</sheetData>");
        readonly byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sheetData>");
        readonly byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
        readonly byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

        private void WriteSharedStrings(Stream stream, Dictionary<string, int> sharedStrings)
        {
            using var writer = new ArrayPoolBufferWriter();
            Encoding.UTF8.GetBytes($@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""{sharedStrings.Count}"">"
                , writer);
            writer.Write(_newLine);
            writer.CopyTo(stream);

            foreach (var s in sharedStrings)
            {
                stream.Write(_siStart);
                Encoding.UTF8.GetBytes(s.Key, writer);
                writer.CopyTo(stream);
                stream.Write(_siEnd);
                stream.Write(_newLine);
            }
            Encoding.UTF8.GetBytes("</sst>", writer);
            writer.Write(_newLine);
            writer.CopyTo(stream);
        }
    }
}
