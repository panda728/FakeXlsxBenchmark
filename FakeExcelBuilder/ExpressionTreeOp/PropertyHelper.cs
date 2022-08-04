using System;
using System.Buffers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FakeExcelBuilder.ExpressionTreeOp
{
    public class PropertyHelper
    {
        public string Name { get; set; } = "";
        public Func<object, object>? Getter { get; set; }
        public int MaxLength { get; set; } = 0;
    }
}
