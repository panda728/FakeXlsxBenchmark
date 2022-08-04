using System.Buffers;
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

            var sheet = new SheetFormatter();
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
                    sheet.CreateSheet(rows, fsSheet, fsString, writeTitle, columnAutoFit);
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
    }
}
