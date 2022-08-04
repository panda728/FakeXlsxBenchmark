using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FakeExcelBuilder
{
    public static class Extentions
    {
        public static void WriteToStream(this string s, Stream output)
            => output.Write(Encoding.UTF8.GetBytes(s));
        public static void WriteToStream(this byte[] bytes, Stream output)
            => output.Write(bytes);
        public static void WriteToStream(this Span<byte> bytes, Stream output)
            => output.Write(bytes);
    }
}
