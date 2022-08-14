﻿using System.Buffers;
using System.Text;

namespace FakeExcel
{
    public static class CellWriter
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

        public static int WriteEmpty(IBufferWriter<byte> writer)
        {
            writer.Write(_emptyColumn);
            return 0;
        }

        public static int Write(string value, IBufferWriter<byte> writer)
        {
            if (string.IsNullOrEmpty(value)) WriteEmpty(writer);

            writer.Write(
                value.Contains(Environment.NewLine)
                    ? _colStartStringWrap
                    : _colStartString
            );

            var index = SharedStrings.TryAdd(value, _index)
                ? _index++
                : SharedStrings[value];
            
            Encoding.UTF8.GetBytes(Convert.ToString(index), writer);

            writer.Write(_colEnd);
            return value.Length;
        }

        static int _index = 0;
        public static Dictionary<string, int> SharedStrings { get; } = new();
        public static void SharedStringsClear()
        {
            _index = 0;
            SharedStrings.Clear();
        }

        public static int Write(object? value, IBufferWriter<byte> writer)
            => Write(value?.ToString() ?? "", writer);
        public static int Write(Guid value, IBufferWriter<byte> writer)
            => Write(value.ToString(), writer);
        public static int Write(Enum value, IBufferWriter<byte> writer)
            => Write(value.ToString(), writer);

        public static int Write(bool? value, IBufferWriter<byte> writer)
        {
            if (value == null) return WriteEmpty(writer);
            var s = Convert.ToString(value);
            if (s == null) return WriteEmpty(writer);

            writer.Write(_colStartBoolean);
            _ = Encoding.UTF8.GetBytes(s, writer);
            writer.Write(_colEnd);
            return s.Length;
        }

        static int WriterNumber(ReadOnlySpan<char> chars, IBufferWriter<byte> writer)
        {
            writer.Write(_colStartNumber);
            _ = Encoding.UTF8.GetBytes(chars, writer);
            writer.Write(_colEnd);
            return chars.Length;
        }

        public static int Write(int value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(long value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(float value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(double value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(decimal value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(short value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);
        public static int Write(ushort value, IBufferWriter<byte> writer)
            => WriterNumber(Convert.ToString(value), writer);

        const int LEN_DATE = 10;
        const int LEN_DATETIME = 18;
        public static int Write(DateTime value, IBufferWriter<byte> writer)
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

        public static int Write(DateOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmpty(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATE;
        }

        public static int Write(TimeOnly? value, IBufferWriter<byte> writer)
        {
            if (value == null) WriteEmpty(writer);
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", writer);
            return LEN_DATETIME;
        }
    }
}