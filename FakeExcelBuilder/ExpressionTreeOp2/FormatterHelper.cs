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

        public static Func<T, IBufferWriter<byte>, long> GenerateFormatter<T>(this PropertyInfo p)
        {
            if (p.PropertyType.IsGenericType)
                return (o, v) => Formatter.WriteEmpty(v);
            if (p.PropertyType == typeof(string))
                return p.GenerateString<T>();
            return IsSupported(p.PropertyType)
                ? p.GenerateSupported<T>()
                : p.GenerateObject<T>();
        }

        static bool IsSupported(Type t)
        {
            if (t.IsPrimitive)
                return true;
            if (t == typeof(string) || t == typeof(Guid) || t == typeof(Enum))
                return true;
            else if (t == typeof(DateTime) || t == typeof(DateOnly) || t == typeof(TimeOnly) || t == _objectType)
                return true;
            return false;
        }

        static Func<T, IBufferWriter<byte>, long> GenerateString<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(Formatter).GetMethod("Write", new Type[] { typeof(ReadOnlySpan<char>), typeof(IBufferWriter<byte>) });
            var methodSpan = typeof(MemoryExtensions).GetMethod("AsSpan", new Type[] { typeof(string) });
            if (method == null || methodSpan == null)
                return (o, v) => Formatter.WriteEmpty(v);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            //var instance = Expression.Convert(target, propertyInfo.DeclaringType);
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var callSpan = Expression.Call(methodSpan, property);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var ps = new Expression[] { callSpan, writer };

            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, writer);
            return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
        }

        static Func<T, IBufferWriter<byte>, long> GenerateSupported<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(Formatter).GetMethod("Serialize", new Type[] { propertyInfo.PropertyType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return (o, v) => Formatter.WriteEmpty(v);

            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            //var instance = Expression.Convert(target, propertyInfo.DeclaringType);
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var ps = new Expression[] { property, writer };

            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, writer);
            return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
        }

        static Func<T, IBufferWriter<byte>, long> GenerateObject<T>(this PropertyInfo propertyInfo)
        {
            var method = typeof(Formatter).GetMethod("Serialize", new Type[] { _objectType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return (o, v) => Formatter.WriteEmpty(v);

            // Func<object, long> getCategoryId = (i,writer) => Formatter.Serialize((object)((i as T).CategoryId), writer);
            var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
            //var instance = Expression.Convert(target, propertyInfo.DeclaringType);
            var property = Expression.PropertyOrField(target, propertyInfo.Name);
            var propertyConv = Expression.Convert(property, _objectType);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");

            var ps = new Expression[] { propertyConv, writer };
            var call = Expression.Call(method, ps);
            var lambda = Expression.Lambda(call, target, writer);
            return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
        }
    }
}
