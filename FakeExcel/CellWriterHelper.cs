using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;

namespace FakeExcel
{
    public class CellWriterHelper<T>
    {
        public CellWriterHelper(string name)
        {
            Name = name;
            Writer = (i, w) => CellWriter.Write(i, w);
            Index = 0;
        }

        public CellWriterHelper(PropertyInfo p, int i)
        {
            Name = p.Name;
            Writer = p.GenerateWriter<T>();
            Index = i;
        }

        public int Index { get; init; }
        public string Name { get; set; }
        public Func<T, IBufferWriter<byte>, int> Writer { get; init; }
    }

    public static class FormatterHelperExtention
    {
        readonly static Type _objectType = typeof(object);
        readonly static Type _bufferWriter = typeof(IBufferWriter<byte>);
        readonly static ParameterExpression _writerParam = Expression.Parameter(_bufferWriter, "w");
        readonly static MethodInfo? _methodObject = typeof(CellWriter).GetMethod("Write", new Type[] { _objectType, _bufferWriter });

        public static Func<T, IBufferWriter<byte>, int> GenerateWriter<T>(this PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType || p.DeclaringType == null)
                return (o, v) => CellWriter.WriteEmpty(v);

            return IsSupported(p.PropertyType)
                ? GenerateSupportedWriter<T>(p.PropertyType, p.DeclaringType, p.Name)
                : GenerateObjectWriter<T>(p.DeclaringType, p.Name);
        }

        static bool IsSupported(Type type)
        {
            if (type.IsPrimitive)
                return true;

            return type == typeof(string)
                || type == typeof(Guid)
                || type == typeof(Enum)
                || type == typeof(DateTime)
                || type == typeof(DateOnly)
                || type == typeof(TimeOnly)
                || type == typeof(object);
        }

        static Func<T, IBufferWriter<byte>, int> GenerateSupportedWriter<T>(
            Type propertyType,
            Type declaringType,
            string name
        )
        {
            var method = typeof(CellWriter).GetMethod("Write", new Type[] { propertyType, _bufferWriter });
            if (method == null)
                return (o, v) => CellWriter.WriteEmpty(v);

            // Func<T, int, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Write(i.CategoryId, writer);
            var target = Expression.Parameter(declaringType, "i");
            var property = Expression.PropertyOrField(target, name);
            var ps = new Expression[] { property, _writerParam };

            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, _writerParam);
            return (Func<T, IBufferWriter<byte>, int>)lambda.Compile();
        }

        static Func<T, IBufferWriter<byte>, int> GenerateObjectWriter<T>(
            Type declaringType,
            string name
        )
        {
            if (_methodObject == null)
                return (o, v) => CellWriter.WriteEmpty(v);

            // Func<T, int, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Write((object)(i.CategoryId), writer);
            var target = Expression.Parameter(declaringType, "i");
            var property = Expression.PropertyOrField(target, name);
            var propertyConv = Expression.Convert(property, _objectType);

            var ps = new Expression[] { propertyConv, _writerParam };
            var call = Expression.Call(_methodObject, ps);
            var lambda = Expression.Lambda(call, target, _writerParam);
            return (Func<T, IBufferWriter<byte>, int>)lambda.Compile();
        }
    }
}
