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
    public static class ColumnFormatter
    {
        //private const int XF_NORMAL = 0;
        private const int XF_WRAP_TEXT = 1;
        private const int XF_DATE = 2;
        private const int XF_DATETIME = 3;

        static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
        static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
        static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@"<c t=""n""><v>");
        static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
        static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
        static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");

        static int _index = 0;
        public static Dictionary<string, int> SharedStrings { get; } = new();

        public static long WriteEmptyCoulumn(IBufferWriter<byte> writer)
        {
            writer.Write(_emptyColumn);
            return 0;
        }

        public static long GetBytes(bool? value, IBufferWriter<byte> writer)
        {
            if (value == null) return WriteEmptyCoulumn(writer);
            writer.Write(_colStartBoolean);
            var s = $"{value}";
            _ = Encoding.UTF8.GetBytes(s, writer);
            writer.Write(_colEnd);
            return s.Length;
        }

        public static long GetBytes(int value, IBufferWriter<byte> writer)
            => WriterNumber($"{value}".AsSpan(), writer);
        public static long GetBytes(long value, IBufferWriter<byte> writer)
            => WriterNumber($"{value}".AsSpan(), writer);
        public static long GetBytes(float value, IBufferWriter<byte> writer)
            => WriterNumber($"{value}".AsSpan(), writer);
        public static long GetBytes(double value, IBufferWriter<byte> writer)
            => WriterNumber($"{value}".AsSpan(), writer);
        public static long GetBytes(decimal value, IBufferWriter<byte> writer)
            => WriterNumber($"{value}".AsSpan(), writer);
        static long WriterNumber(ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
        {
            writer.Write(_colStartNumber);
            _ = Encoding.UTF8.GetBytes(chars, writer);
            writer.Write(_colEnd);
            return chars.Length;
        }

        private const int LEN_DATE = 10;
        private const int LEN_DATETIME = 18;
        public static long GetBytes(DateTime value, IBufferWriter<byte> writer)
        {
            var d = value;
            if (d == DateTime.MinValue) WriteEmptyCoulumn(writer);
            if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
            {
                Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
                return LEN_DATE;
            }
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATETIME;
        }
        public static long GetBytes(DateOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmptyCoulumn(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATE;
        }
        public static long GetBytes(TimeOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmptyCoulumn(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATETIME;
        }

        public static long GetBytes(object value, IBufferWriter<byte> writer)
            => GetBytes(value?.ToString() ?? "", writer);

        public static long GetBytes(Guid value, IBufferWriter<byte> writer)
        {
            var s = value.ToString();
            GetBytes(s, writer);
            return s.Length;
        }

        public static long GetBytes(Enum value, IBufferWriter<byte> writer)
        {
            var s = value.ToString();
            GetBytes(s, writer);
            return s.Length;
        }

        public static long GetBytes(string value, IBufferWriter<byte> writer)
        {
            if (string.IsNullOrEmpty(value)) WriteEmptyCoulumn(writer);

            var index = GetSharedStringIndex(value);
            if (value.Contains(Environment.NewLine))
                writer.Write(_colStartStringWrap);
            else
                writer.Write(_colStartString);

            Encoding.ASCII.GetBytes($"{index}",writer);
            writer.Write(_colEnd);
            return value.Length;
        }

        public static void SharedStringsClear()
        {
            _index = 0;
            SharedStrings.Clear();
        }

        static int GetSharedStringIndex(string s)
        {
            if (SharedStrings.ContainsKey(s))
                return SharedStrings[s];

            SharedStrings.Add(s, _index);
            return _index++;
        }
    }
}

