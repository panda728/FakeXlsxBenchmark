using System.Buffers;
using System.Text;

namespace FakeExcelBuilder.ExpressionTreeOp2
{
    public static class Formatter
    {
        //private const int XF_NORMAL = 0;
        const int XF_WRAP_TEXT = 1;
        const int XF_DATE = 2;
        const int XF_DATETIME = 3;

        static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
        static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
        static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@"<c t=""n""><v>");
        static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
        static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
        static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");

        public static long WriteEmpty(IBufferWriter<byte> writer)
        {
            writer.Write(_emptyColumn);
            return 0;
        }

        #region string
        static int _index = 0;
        public static Dictionary<string, int> SharedStrings { get; } = new();

        public static long Serialize(string value, IBufferWriter<byte> writer)
        {
            if (string.IsNullOrEmpty(value)) WriteEmpty(writer);

            var index = GetSharedStringIndex(value);
            if (value.Contains(Environment.NewLine))
                writer.Write(_colStartStringWrap);
            else
                writer.Write(_colStartString);

            Encoding.UTF8.GetBytes(Convert.ToString(index), writer);
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
        #endregion

        public static long Serialize(object value, IBufferWriter<byte> writer)
            => Serialize(value?.ToString() ?? "", writer);
        public static long Serialize(Guid value, IBufferWriter<byte> writer)
            => Serialize(value.ToString(), writer);
        public static long Serialize(Enum value, IBufferWriter<byte> writer)
            => Serialize(value.ToString(), writer);

        #region number
        public static long Serialize(int value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static long Serialize(long value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static long Serialize(float value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static long Serialize(double value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static long Serialize(decimal value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);

        static long WriterNumber(ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
        {
            writer.Write(_colStartNumber);
            _ = Encoding.UTF8.GetBytes(chars, writer);
            writer.Write(_colEnd);
            return chars.Length;
        }
        #endregion

        public static long Serialize(bool? value, IBufferWriter<byte> writer)
        {
            if (value == null) return WriteEmpty(writer);
            writer.Write(_colStartBoolean);
            var s = Convert.ToString(value);
            _ = Encoding.UTF8.GetBytes(s, writer);
            writer.Write(_colEnd);
            return s.Length;
        }

        #region date
        const int LEN_DATE = 10;
        const int LEN_DATETIME = 18;
        public static long Serialize(DateTime value, IBufferWriter<byte> writer)
        {
            var d = value;
            if (d == DateTime.MinValue) WriteEmpty(writer);
            if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
            {
                Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
                return LEN_DATE;
            }
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATETIME;
        }
        public static long Serialize(DateOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmpty(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATE;
        }
        public static long Serialize(TimeOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmpty(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATETIME;
        }
        #endregion
    }
}

