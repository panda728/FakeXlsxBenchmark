using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;

namespace FakeExcelBuilder.ExpressionTreeOp2;

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
    public SerializeDelegate<T> Formatter { get; init; }
    public int MaxLength { get; set; }
}
public delegate long SerializeDelegate<T>(T value, ref IBufferWriter<byte> writer);
public static class FormatterHelperExtention
{
    readonly static Type _objectType = typeof(object);
    readonly static (MethodInfo method, Type? type)[] _methods =
        typeof(Formatter)
            .GetMethods()
            .Where(m => m.Name == "Serialize")
            .Select(m => (m, m.GetParameters()?.FirstOrDefault()?.ParameterType))
            .ToArray();

    readonly static MethodInfo _methodObject = _methods.Where(x => x.type == typeof(object)).First().method;
    readonly static ParameterExpression _writer = Expression.Parameter(typeof(IBufferWriter<byte>).MakeByRefType(), "w");

    public static SerializeDelegate<T> GenerateFormatter<T>(this PropertyInfo p)
    {
        if (p.DeclaringType == null || p.PropertyType.IsGenericType)
            return (T value, ref IBufferWriter<byte> writer) => Formatter.WriteEmpty(ref writer);

        var methodTyped = _methods.Where(x => x.type == p.PropertyType)?.Select(x => x.method)?.FirstOrDefault();
        var target = Expression.Parameter(p.DeclaringType, "i");
        var parameters = methodTyped == null
            ? new Expression[] { Expression.Convert(Expression.PropertyOrField(target, p.Name), _objectType), _writer }
            : new Expression[] { Expression.PropertyOrField(target, p.Name), _writer };

        var call = Expression.Call(methodTyped ?? _methodObject, parameters);
        var lambda = Expression.Lambda<SerializeDelegate<T>>(call, target, _writer);
        return lambda.Compile();
    }
}

//public static class FormatterHelperExtention
//{
//    readonly static Type _objectType = typeof(object);
//    readonly static Type _bufferWriter = typeof(IBufferWriter<byte>);
//    readonly static ParameterExpression _writerParam = Expression.Parameter(_bufferWriter, "w");
//    readonly static MethodInfo? _methodObject = typeof(Formatter).GetMethod("Serialize", new Type[] { _objectType, _bufferWriter });

//    public static Func<T, IBufferWriter<byte>, long> GenerateFormatter<T>(this PropertyInfo p)
//    {
//        if (p.PropertyType.IsGenericType)
//            return (o, v) => Formatter.WriteEmpty(v);
//        return IsSupported(p.PropertyType)
//            ? p.GenerateSupported<T>()
//            : p.GenerateObject<T>();
//    }

//    static bool IsSupported(Type type)
//    {
//        if (type.IsPrimitive)
//            return true;

//        return type == typeof(string)
//            || type == typeof(Guid)
//            || type == typeof(Enum)
//            || type == typeof(DateTime)
//            || type == typeof(DateOnly)
//            || type == typeof(TimeOnly)
//            || type == typeof(object);
//    }

//    static Func<T, IBufferWriter<byte>, long> GenerateSupported<T>(this PropertyInfo propertyInfo)
//    {
//        var method = typeof(Formatter).GetMethod("Serialize", new Type[] { propertyInfo.PropertyType, _bufferWriter });
//        if (method == null || propertyInfo.DeclaringType == null)
//            return (o, v) => Formatter.WriteEmpty(v);

//        // Func<T, long, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Serialize(i.CategoryId, writer);
//        var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
//        var property = Expression.PropertyOrField(target, propertyInfo.Name);
//        var ps = new Expression[] { property, _writerParam };

//        var call = Expression.Call(method, ps);
//        var lambda = Expression.Lambda(call, target, _writerParam);
//        return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
//    }

//    static Func<T, IBufferWriter<byte>, long> GenerateObject<T>(this PropertyInfo propertyInfo)
//    {
//        if (_methodObject == null || propertyInfo.DeclaringType == null)
//            return (o, v) => Formatter.WriteEmpty(v);

//        // Func<T, long, IBufferWriter<byte>> getCategoryId = (i,writer) => Formatter.Serialize((object)(i.CategoryId), writer);
//        var target = Expression.Parameter(propertyInfo.DeclaringType, "i");
//        var property = Expression.PropertyOrField(target, propertyInfo.Name);
//        var propertyConv = Expression.Convert(property, _objectType);

//        var ps = new Expression[] { propertyConv, _writerParam };
//        var call = Expression.Call(_methodObject, ps);
//        var lambda = Expression.Lambda(call, target, _writerParam);
//        return (Func<T, IBufferWriter<byte>, long>)lambda.Compile();
//    }
//}
