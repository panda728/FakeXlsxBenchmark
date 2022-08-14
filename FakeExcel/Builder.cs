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
        readonly byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
        readonly byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

        readonly byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets>
<sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/>
</sheets>
</workbook>");
        readonly byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

        readonly byte[] _styles = Encoding.UTF8.GetBytes(@"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
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

        const int COLUMN_WIDTH_MAX = 100;
        const int COLUMN_WIDTH_MARGIN = 2;

        static class GetPropertiesCache<T>
        {
            static GetPropertiesCache()
            {
                var type = typeof(T);
                var name = type.FullName?.Split(".")?.FirstOrDefault() ?? "";
                if (name == "System")
                {
                    Properties = new CellWriterHelper<T>[] { new CellWriterHelper<T>("value") };
                    return;
                }

                Properties = type
                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .AsParallel()
                .Select((p, i) => new CellWriterHelper<T>(p, i))
                .OrderBy(p => p.Index)
                .ToArray();
            }
            public static readonly CellWriterHelper<T>[] Properties;
        }

        public void CreateExcelFile<T>(
            string fileName,
            IEnumerable<T> rows,
            bool showTitleRow = false,
            bool columnAutoFit = false,
            string[]? titles = null,
            string workPath = "work"
        )
        {
            var workPathRoot = Path.Combine(workPath, Guid.NewGuid().ToString());
            try
            {
                if (!Directory.Exists(workPathRoot))
                    Directory.CreateDirectory(workPathRoot);

                using (var sheetStream = CreateStream(Path.Combine(workPathRoot, "sheet.xml")))
                using (var stringsStream = CreateStream(Path.Combine(workPathRoot, "strings.xml")))
                {
                    CellWriter.SharedStringsClear();
                    CreateSheet(rows, sheetStream, showTitleRow, columnAutoFit, titles ?? Array.Empty<string>());
                    WriteSharedStrings(stringsStream);
                    CellWriter.SharedStringsClear();
                }

                var workRelPath = Path.Combine(workPathRoot, "_rels");
                if (!Directory.Exists(workRelPath))
                    Directory.CreateDirectory(workRelPath);

                Task.WaitAll(
                    new Task[] {
                        WriteStreamAsync(_contentTypes, Path.Combine(workPathRoot, "[Content_Types].xml")),
                        WriteStreamAsync(_book, Path.Combine(workPathRoot, "book.xml")),
                        WriteStreamAsync(_styles, Path.Combine(workPathRoot, "styles.xml")),
                        WriteStreamAsync(_rels, Path.Combine(workRelPath, ".rels")),
                        WriteStreamAsync(_bookRels, Path.Combine(workRelPath, "book.xml.rels"))
                    });

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

        async Task WriteStreamAsync(ReadOnlyMemory<byte> bytes, string fileName)
        {
            using (var fs = CreateStream(fileName))
                await fs.WriteAsync(bytes);
        }

        Stream CreateStream(string fileName)
            => new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);

        void CreateSheet<T>(
            IEnumerable<T> rows,
            Stream stream,
            bool showTitleRow,
            bool autoFitColumns,
            string[] titles
        )
        {
            stream.Write(_sheetStart);

            if (showTitleRow)
                stream.Write(_frozenTitleRow);

            var formatters = GetPropertiesCache<T>.Properties.AsSpan();
            if (autoFitColumns)
                WriteCellWidth(formatters, rows, stream);

            stream.Write(_dataStart);

            using var writer = new ArrayPoolBufferWriter<byte>();
            if (showTitleRow)
            {
                stream.Write(_rowStart);
                foreach (var f in formatters)
                {
                    CellWriter.Write(
                        titles.Length > f.Index ? titles[f.Index] : f.Name,
                        writer
                    );
                    stream.Write(writer.WrittenSpan);
                    writer.Clear();
                }
                stream.Write(_rowEnd);
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
            }

            stream.Write(_dataEnd);
            stream.Write(_sheetEnd);
        }

        void WriteCellWidth<T>(
            Span<CellWriterHelper<T>> formatters,
            IEnumerable<T> rows,
            Stream stream
        )
        {
            var i = 0;
            stream.Write(_colStart);

            using var writer = new ArrayPoolBufferWriter<byte>();
            foreach (var f in formatters)
            {
                ++i;
                var maxLength = GetMaxLength(f, rows, writer);
                Encoding.UTF8.GetBytes(
                    @$"<col min=""{i}"" max =""{i}"" width =""{maxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
                    writer);
                stream.Write(writer.WrittenSpan);
                writer.Clear();
            }
            stream.Write(_colEnd);
        }

        int GetMaxLength<T>(
            CellWriterHelper<T> f,
            IEnumerable<T> rows,
            ArrayPoolBufferWriter<byte> writer
        )
        {
            var max = rows
                .Take(100)
                .Select(r =>
                {
                    var len = f?.Writer(r, writer) ?? 0;
                    writer.Clear();
                    return len;
                })
                .Max(x => x);

            var lenMax = Math.Max(max, f.Name.Length) + COLUMN_WIDTH_MARGIN;
            return Math.Min(lenMax, COLUMN_WIDTH_MAX);
        }

        void WriteSharedStrings(Stream stream)
        {
            stream.Write(_sstStart);

            using var writer = new ArrayPoolBufferWriter<byte>();
            foreach (var s in CellWriter.SharedStrings.Keys)
            {
                stream.Write(_siStart);

                Encoding.UTF8.GetBytes(s, writer);
                stream.Write(writer.WrittenSpan);
                writer.Clear();

                stream.Write(_siEnd);
            }

            stream.Write(_sstEnd);
        }
    }
}
