using System.Buffers;
using System.Collections.Concurrent;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using Microsoft.Toolkit.HighPerformance.Buffers;

namespace FakeExcel
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

        readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
        //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
        readonly byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
        readonly byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
        readonly byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

        private const int COLUMN_WIDTH_MAX = 100;
        private const int COLUMN_WIDTH_MARGIN = 2;

        static class GetPropertiesCache<T>
        {
            static GetPropertiesCache()
            {
                var type = typeof(T);
                if (type.Namespace?.StartsWith("System") ?? true)
                {
                    Properties = new FormatterHelper<T>[] { new FormatterHelper<T>("value") };
                    return;
                }

                Properties = typeof(T)
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                    .AsParallel()
                    .Select((p, i) => new FormatterHelper<T>(p, i))
                    .OrderBy(p => p.Index)
                    .ToArray();
            }
            public static readonly FormatterHelper<T>[] Properties;
        }

        public async Task CreateExcelFileAsync<T>(
            string fileName,
            IEnumerable<T> rows,
            bool showTitleRow = false,
            bool columnAutoFit = false,
            string[]? titles = null,
            string workPath = "work"
        )
        {
            var formatters = GetPropertiesCache<T>.Properties;
            if (titles != null)
            {
                var i = 0;
                foreach (var f in formatters)
                {
                    if (titles.Length > i)
                        f.Name = titles[i++];
                }
            }

            var workPathRoot = Path.Combine(workPath, Guid.NewGuid().ToString());
            try
            {
                if (!Directory.Exists(workPathRoot))
                    Directory.CreateDirectory(workPathRoot);

                var workRelPath = Path.Combine(workPathRoot, "_rels");
                if (!Directory.Exists(workRelPath))
                    Directory.CreateDirectory(workRelPath);

                using (var fs = CreateStream(Path.Combine(workPathRoot, "[Content_Types].xml")))
                    await fs.WriteAsync(_contentTypes);
                using (var fs = CreateStream(Path.Combine(workRelPath, ".rels")))
                    await fs.WriteAsync(_rels);
                using (var fs = CreateStream(Path.Combine(workPathRoot, "book.xml")))
                    await fs.WriteAsync(_book);
                using (var fs = CreateStream(Path.Combine(workRelPath, "book.xml.rels")))
                    await fs.WriteAsync(_bookRels);
                using (var fs = CreateStream(Path.Combine(workPathRoot, "styles.xml")))
                    await fs.WriteAsync(_styles);
                using (var stream = CreateStream(Path.Combine(workPathRoot, "sheet.xml")))
                using (var fsString = CreateStream(Path.Combine(workPathRoot, "strings.xml")))
                {
                    Formatter.SharedStringsClear();
                    CreateSheet(formatters.AsSpan(), rows, stream, showTitleRow, columnAutoFit);
                    WriteSharedStrings(fsString, Formatter.SharedStrings);
                    Formatter.SharedStringsClear();
                }

                if (File.Exists(fileName))
                    File.Delete(fileName);
                ZipFile.CreateFromDirectory(workPathRoot, fileName);
            }
            catch
            {
                throw;
            }
            finally
            {
                try
                {
                    if (Directory.Exists(workPathRoot))
                        Directory.Delete(workPathRoot, true);
                }
                catch { }
            }
        }

        Stream CreateStream(string fileName)
            => new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);

        void CalcCellStringLength<T>(Span<FormatterHelper<T>> formatters, IEnumerable<T> rows)
        {
            using var buffer = new ArrayPoolBufferWriter<byte>();
            foreach (var f in formatters)
            {
                var maxLength = rows
                    .Take(100)
                    .Select(r =>
                    {
                        if (r == null || f.Writer == null) return 0;
                        var len = (int)f.Writer(r, buffer);
                        buffer.Clear();
                        return len;
                    })
                    .Max(x => x);
                f.MaxLength = Math.Min(
                    Math.Max(maxLength, f.Name.Length) + COLUMN_WIDTH_MARGIN,
                    COLUMN_WIDTH_MAX);
            }
        }

        void CreateSheet<T>(Span<FormatterHelper<T>> formatters, IEnumerable<T> rows, Stream stream, bool showTitleRow, bool autoFitColumns)
        {
            stream.Write(_sheetStart);
            stream.Write(_newLine);

            if (showTitleRow)
            {
                stream.Write(_frozenTitleRow);
                stream.Write(_newLine);
            }

            using var writer = new ArrayPoolBufferWriter<byte>();
            if (autoFitColumns)
            {
                CalcCellStringLength(formatters, rows);
                var i = 0;
                stream.Write(_colStart);
                foreach (var f in formatters)
                {
                    ++i;
                    Encoding.UTF8.GetBytes(
                        @$"<col min=""{i}"" max =""{i}"" width =""{f.MaxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
                        writer);
                    stream.Write(writer.WrittenSpan);
                    writer.Clear();
                }
                stream.Write(_colEnd);
                stream.Write(_newLine);
            }

            stream.Write(_dataStart);
            stream.Write(_newLine);

            if (showTitleRow)
            {
                stream.Write(_rowStart);
                foreach (var f in formatters)
                {
                    Formatter.Write(f.Name, writer);
                    stream.Write(writer.WrittenSpan);
                    writer.Clear();
                }
                stream.Write(_rowEnd);
                stream.Write(_newLine);
            }

            foreach (var row in rows)
            {
                if (row == null) continue;
                stream.Write(_rowStart);
                foreach (var f in formatters)
                {
                    f.Writer(row, writer);
                    stream.Write(writer.WrittenSpan);
                    writer.Clear();
                }
                stream.Write(_rowEnd);
                stream.Write(_newLine);
            }
            stream.Write(_dataEnd);
            stream.Write(_newLine);
            stream.Write(_sheetEnd);
            stream.Write(_newLine);
        }

        void WriteSharedStrings(Stream stream, Dictionary<string, int> sharedStrings)
        {
            stream.Write(_sstStart);
            stream.Write(_newLine);

            using var writer = new ArrayPoolBufferWriter<byte>();
            foreach (var s in sharedStrings)
            {
                writer.Write(_siStart);
                Encoding.UTF8.GetBytes(s.Key, writer);
                writer.Write(_siEnd);
                writer.Write(_newLine);
                stream.Write(writer.WrittenSpan);
                writer.Clear();
            }
            stream.Write(_sstEnd);
            stream.Write(_newLine);
        }
    }
}
