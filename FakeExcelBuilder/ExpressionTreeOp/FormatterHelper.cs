using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;

namespace FakeExcelBuilder.ExpressionTreeOp
{
    public class FormatterHelper
    {
        public FormatterHelper(Type t, PropertyInfo p, int i)
        {
            Name = p.Name;
            Formatter = FormatterHelperExtention.GenerateEncodedGetterLambda(t, p);
            MaxLength = 0;
            Index = i;
        }

        /// <summary>
        /// index
        /// </summary>
        public int Index { get; init; }
        /// <summary>
        /// for title
        /// </summary>
        public string Name { get; init; }
        /// <summary>
        /// Build Column Xml
        /// </summary>
        public Func<object, IBufferWriter<byte>, long> Formatter { get; init; }
        /// <summary>
        /// for autofit
        /// </summary>
        public int MaxLength { get; set; }
    }

    public static class FormatterHelperExtention
    {
        readonly static Type _objectType = typeof(object);

        public static Func<object, IBufferWriter<byte>, long> GenerateEncodedGetterLambda(Type t, PropertyInfo p)
            => IsSupported(p.PropertyType)
                ? p.GeneratePrimitiveLambda(t)
                : p.GenerateObjectLambda(t);

        static bool IsSupported(Type t)
        {
            if (t == typeof(string) || t == typeof(Guid) || t == typeof(Enum))
                return true;
            else if (t == typeof(int) || t == typeof(long) || t == typeof(decimal) || t == typeof(double) || t == typeof(float))
                return true;
            else if (t == typeof(DateTime) || t == typeof(DateOnly) || t == typeof(TimeOnly) || t == _objectType)
                return true;
            return false;
        }

        static Func<object, IBufferWriter<byte>, long> GeneratePrimitiveLambda(this PropertyInfo propertyInfo, Type t)
        {
            if (t != propertyInfo.DeclaringType || propertyInfo.PropertyType.IsGenericType || propertyInfo.PropertyType.IsArray)
                return (o, v) => ColumnFormatter.WriteEmptyCoulumn(o, v);

            // Build an equivalent process with ExpresstionTree.
            // Func<T, long> getCategoryId = (i,writer) => Formatter.GetBytes((i as T).CategoryId,writer);
            var instanceObj = Expression.Parameter(_objectType, "i");
            var instance = Expression.Convert(instanceObj, propertyInfo.DeclaringType);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var property = Expression.Property(instance, propertyInfo);
            var ps = new Expression[] { property, writer };
            var method = typeof(ColumnFormatter).GetMethod("Serialize", new Type[] { propertyInfo.PropertyType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return (o, v) => ColumnFormatter.WriteEmptyCoulumn(o, v);

            var call = Expression.Call(method, ps);
            var d = Expression.Lambda(call, instanceObj, writer).Compile();
            return (Func<object, IBufferWriter<byte>, long>)d;
        }

        static Func<object, IBufferWriter<byte>, long> GenerateObjectLambda(this PropertyInfo propertyInfo, Type t)
        {
            if (t != propertyInfo.DeclaringType || propertyInfo.PropertyType.IsGenericType || propertyInfo.PropertyType.IsArray)
                return (o, v) => ColumnFormatter.WriteEmptyCoulumn(o, v);

            // Build an equivalent process with ExpresstionTree.
            // Func<T, long> getCategoryId = (i,writer) => Formatter.GetBytes((object)((i as T).CategoryId),writer);
            var instanceObj = Expression.Parameter(_objectType, "i");
            var instance = Expression.Convert(instanceObj, propertyInfo.DeclaringType);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var property = Expression.Property(instance, propertyInfo);
            var propertyObject = Expression.Convert(property, _objectType);
            var ps = new Expression[] { propertyObject, writer };
            var method = typeof(ColumnFormatter).GetMethod("Serialize", new Type[] { _objectType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return (o, v) => ColumnFormatter.WriteEmptyCoulumn(o, v);

            var call = Expression.Call(method, ps);
            var d = Expression.Lambda(call, instanceObj, writer).Compile();
            return (Func<object, IBufferWriter<byte>, long>)d;
        }
    }
}
