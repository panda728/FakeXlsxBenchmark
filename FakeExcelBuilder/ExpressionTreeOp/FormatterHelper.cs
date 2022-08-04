using System;
using System.Buffers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FakeExcelBuilder.ExpressionTreeOp
{
    public class FormatterHelper<T>
    {
        public string Name { get; set; } = "";
        public Func<T, IBufferWriter<byte>, long>? Formatter { get; set; }
    }
}
