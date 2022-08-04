using System;
using System.Buffers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FakeExcelBuilder.ExpressionTreeOp
{
    public class FormatterHelper
    {
        public string Name { get; set; } = "";
        public Func<object, IBufferWriter<byte>, long>? Formatter { get; set; }
        public int MaxLength { get; set; } = 0;
    }
}
