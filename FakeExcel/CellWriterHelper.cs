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

        public static Func<T, IBufferWriter<byte>, int> GenerateWriter<T>(this PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType)
                return (o, v) => CellWriter.WriteEmpty(v);
            return IsSupported(p.PropertyType)
                ? p.GenerateSupportedWriter<T>()
                : p.GenerateObjectWriter<T>();
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

        static Func<T, IBufferWriter<byte>, int> GenerateSupportedWriter<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(CellWriter).GetMethod("Write", new Type[] { propertyInfo.PropertyType, _bufferWriter });
            if (method == null || propertyInfo.DeclaringType == null)
                return (o, v) => CellWriter.WriteEmpty(v);

            // Func<T, int, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Write(i.CategoryId, writer);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var ps = new Expression[] { property, _writerParam };

            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, _writerParam);
            return (Func<T, IBufferWriter<byte>, int>)lambda.Compile();
        }

        static Func<T, IBufferWriter<byte>, int> GenerateObjectWriter<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(CellWriter).GetMethod("Write", new Type[] { _objectType, _bufferWriter });
            if (method == null || propertyInfo.DeclaringType == null)
                return (o, v) => CellWriter.WriteEmpty(v);

            // Func<T, int, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Write((object)(i.CategoryId), writer);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var propertyConv = Expression.Convert(property, _objectType);

            var ps = new Expression[] { propertyConv, _writerParam };
            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, _writerParam);
            return (Func<T, IBufferWriter<byte>, int>)lambda.Compile();
        }
    }
}
