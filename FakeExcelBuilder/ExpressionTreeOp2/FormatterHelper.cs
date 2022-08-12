using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;

namespace FakeExcelBuilder.ExpressionTreeOp2
{
    public class FormatterHelper<T>
    {
        public FormatterHelper(PropertyInfo p, int i)
        {
            Name = p.Name;
            Formatter = p.GenerateFormatter<T>();
            Index = i;
            MaxLength = 0;
        }

        public int Index { get; init; }
        public string Name { get; init; }
        public Func<T, IBufferWriter<byte>, long> Formatter { get; init; }
        public int MaxLength { get; set; }
    }

    public static class FormatterHelperExtention
    {
        readonly static Type _objectType = typeof(object);
        readonly static Type _bufferWriter = typeof(IBufferWriter<byte>);

        public static Func<T, IBufferWriter<byte>, long> GenerateFormatter<T>(this PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType)
                return (o, v) => Formatter.WriteEmpty(v);
            return IsSupported(p.PropertyType)
                ? p.GenerateSupported<T>()
                : p.GenerateObject<T>();
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

        static Func<T, IBufferWriter<byte>, long> GenerateSupported<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(Formatter).GetMethod("Serialize", new Type[] { propertyInfo.PropertyType, _bufferWriter });
            if (method == null || propertyInfo.DeclaringType == null)
                return (o, v) => Formatter.WriteEmpty(v);

            // Func<T, long, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Serialize(i.CategoryId, writer);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var writer = Expression.Parameter(_bufferWriter, "w");
            var ps = new Expression[] { property, writer };

            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, writer);
            return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
        }

        static Func<T, IBufferWriter<byte>, long> GenerateObject<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(Formatter).GetMethod("Serialize", new Type[] { _objectType, _bufferWriter });
            if (method == null || propertyInfo.DeclaringType == null)
                return (o, v) => Formatter.WriteEmpty(v);

            // Func<T, long, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Serialize((object)(i.CategoryId), writer);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var propertyConv = Expression.Convert(property, _objectType);
            var writer = Expression.Parameter(_bufferWriter, "w");

            var ps = new Expression[] { propertyConv, writer };
            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, writer);
            return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
        }
    }
}
