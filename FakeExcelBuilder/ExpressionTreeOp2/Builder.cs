using Microsoft.Toolkit.HighPerformance.Buffers;
using System.Buffers;
using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace FakeExcelBuilder.ExpressionTreeOp2;

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

    readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
    //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
    readonly byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
    readonly byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
    readonly byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

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

    public async Task RunAsync<T>(string fileName, IEnumerable<T> rows, bool showTitleRow = true, bool columnAutoFit = true)
    {
        var workPath = Path.Combine("work", Guid.NewGuid().ToString());
        var workRelPath = Path.Combine(workPath, "_rels");
#if DEBUG
        if (!Directory.Exists(workPath))
            Directory.CreateDirectory(workPath);

        if (!Directory.Exists(workRelPath))
            Directory.CreateDirectory(workRelPath);
#endif
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
                CreateSheet(rows, fsSheet, showTitleRow, columnAutoFit);
                WriteSharedStrings(fsString);
                Formatter.SharedStringsClear();
            }
#if DEBUG
            if (File.Exists(fileName))
                File.Delete(fileName);
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
            //try
            //{
            //    Directory.Delete(workPath, true);
            //}
            //catch { }
#else
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

    void WriteCellWidth<T>(Span<FormatterHelper<T>> formatters, IEnumerable<T> rows, Stream fsSheet)
    {
        var i = 0;
        fsSheet.Write(_colStart);

        using var buffer = new ArrayPoolBufferWriter<byte>();
        var writer = (IBufferWriter<byte>)buffer;

        foreach (var f in formatters)
        {
            var max = rows
                .Take(100)
                .Select(r =>
                {
                    var len = f?.Formatter(r, ref writer) ?? 0;
                    buffer.Clear();
                    return len;
                })
                .Max(x => x);

            var lenMax = (int)Math.Max(max, f.Name.Length) + COLUMN_WIDTH_MARGIN;
            var maxLength = Math.Min(lenMax, COLUMN_WIDTH_MAX);
            ++i;
            Encoding.UTF8.GetBytes(
                @$"<col min=""{i}"" max =""{i}"" width =""{maxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
                buffer);
            fsSheet.Write(buffer.WrittenSpan);
            buffer.Clear();
            fsSheet.Write(_newLine);
        }
        fsSheet.Write(_colEnd);
        fsSheet.Write(_newLine);
    }

    public void CreateSheet<T>(IEnumerable<T> rows, Stream fsSheet, bool showTitleRow, bool autoFitColumns)
    {
        fsSheet.Write(_sheetStart);
        fsSheet.Write(_newLine);

        if (showTitleRow)
        {
            fsSheet.Write(_frozenTitleRow);
            fsSheet.Write(_newLine);
        }

        var formatters = GetPropertiesCache<T>.Properties.AsSpan();
        if (autoFitColumns)
            WriteCellWidth(formatters, rows, fsSheet);

        fsSheet.Write(_dataStart);
        fsSheet.Write(_newLine);

        using var buffer = new ArrayPoolBufferWriter<byte>();
        var writer = (IBufferWriter<byte>)buffer;
        if (showTitleRow)
        {
            fsSheet.Write(_rowStart);
            foreach (var f in formatters)
            {
                Formatter.Serialize(f.Name, ref writer);
                fsSheet.Write(buffer.WrittenSpan);
                buffer.Clear();
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
                f.Formatter(row, ref writer);
            }
            fsSheet.Write(buffer.WrittenSpan);
            buffer.Clear();
            fsSheet.Write(_rowEnd);
            fsSheet.Write(_newLine);
        }
        fsSheet.Write(_dataEnd);
        fsSheet.Write(_newLine);
        fsSheet.Write(_sheetEnd);
        fsSheet.Write(_newLine);
    }

    private void WriteSharedStrings(Stream stream)
    {
        stream.Write(_sstStart);
        stream.Write(_newLine);

        using var buffer = new ArrayBufferWriter();
        foreach (var s in Formatter.SharedStrings.Keys)
        {
            stream.Write(_siStart);
            Encoding.UTF8.GetBytes(s, buffer);
            buffer.CopyTo(stream);
            stream.Write(_siEnd);
            stream.Write(_newLine);
        }
        stream.Write(_sstEnd);
        stream.Write(_newLine);
    }
}
