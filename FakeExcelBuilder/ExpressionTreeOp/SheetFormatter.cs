using System;
using System.Buffers;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;


namespace FakeExcelBuilder.ExpressionTreeOp
{
    public class SheetFormatter
    {
        readonly ConcurrentDictionary<Type, PropertyHelper[]> _properties = new();

        readonly byte[] _newLine = Encoding.UTF8.GetBytes(Environment.NewLine);
        readonly byte[] _rowTag1 = Encoding.UTF8.GetBytes("<row>");
        readonly byte[] _rowTag2 = Encoding.UTF8.GetBytes("</row>");
        readonly byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

        public void CreateSheet<T>(IEnumerable<T> rows, Stream fsSheet, Stream fsString, bool writeTitle, bool autoFitColumns)
        {
            ColumnFormatter<T>.SharedStringsClear();
            var formatters = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(x => new FormatterHelper<T>()
                {
                    Name = x.Name,
                    Formatter = ColumnFormatter<T>.IsSupportedType(x.PropertyType)
                        ? PropertyInfoExtensions.GenerateEncodedGetterLambda<T>(x)
                        : null
                })
                .ToArray().AsSpan();

            var properties = (writeTitle || autoFitColumns)
                ? _properties.GetOrAdd(typeof(T), typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(x => new PropertyHelper()
                {
                    Name = x.Name,
                    Getter = PropertyInfoExtensions.GenerateGetterLambda<T>(x)
                })
                .ToArray())
                : null;

            if (autoFitColumns)
            {
                foreach (var p in properties.AsSpan())
                {
                    var maxLength = rows
                        .Take(100)
                        .Select(r =>
                        {
                            if (r == null || p.Getter == null) return "";
                            return p.Getter(r)?.ToString() ?? "";
                        })
                        .Max(s => s.Length);
                    p.MaxLength = Math.Min(
                        Math.Max(maxLength, p.Name.Length) + 2,
                        100);
                }
            }

            WriteHeader(fsSheet, properties, writeTitle, autoFitColumns);

            var writer = new ArrayBufferWriter<byte>();
            writer.Write(Encoding.UTF8.GetBytes("<sheetData>"));
            writer.Write(_newLine);

            if (writeTitle)
            {
                writer.Write(_rowTag1);
                foreach (var p in properties.AsSpan())
                    ColumnFormatter<T>.GetBytes(p.Name, writer);
                writer.Write(_rowTag2);
                writer.Write(_newLine);
            }
            fsSheet.Write(writer.WrittenSpan);
            writer.Clear();

            foreach (var row in rows)
            {
                if (row == null) continue;
                writer.Write(_rowTag1);
                foreach (var f in formatters)
                {
                    if (f.Formatter == null)
                        _ = ColumnFormatter<T>.WriteEmptyCoulumn(writer);
                    else
                        _ = f.Formatter(row, writer);
                }
                writer.Write(_rowTag2);
                writer.Write(_newLine);
                fsSheet.Write(writer.WrittenSpan);
                writer.Clear();
            }

            fsSheet.Write(Encoding.UTF8.GetBytes("</sheetData>"));
            fsSheet.Write(_newLine);
            fsSheet.Write(Encoding.UTF8.GetBytes("</worksheet>"));
            fsSheet.Write(_newLine);

            WriteSharedStrings(fsString, ColumnFormatter<T>.SharedStrings);
            ColumnFormatter<T>.SharedStringsClear();
        }

        void WriteHeader(Stream stream, PropertyHelper[] properties, bool writeTitle, bool autoFitColumns)
        {
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

            if (autoFitColumns && properties != null)
            {
                var i = 0;
                "<cols>".WriteToStream(stream);
                foreach (var p in properties.AsSpan())
                {
                    ++i;
                    @$"<col min=""{i}"" max =""{i}"" width =""{p.MaxLength:0.0}"" bestFit =""1"" customWidth =""1"" />"
                        .WriteToStream(stream);
                }
                "</cols>".WriteToStream(stream);
                _newLine.WriteToStream(stream);
            }
        }

        private void WriteSharedStrings(Stream stream, Dictionary<string, int> sharedStrings)
        {
            $@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""{sharedStrings.Count}"">"
                .WriteToStream(stream);
            _newLine.WriteToStream(stream);

            var tagA = Encoding.UTF8.GetBytes("<si><t>").AsSpan();
            var tagB = Encoding.UTF8.GetBytes("</t></si>").AsSpan();
            foreach (var s in sharedStrings)
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
