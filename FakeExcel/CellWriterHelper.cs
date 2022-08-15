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
            Writer = (T value, ref IBufferWriter<byte> writer) => CellWriter.Write(value, ref writer);
            Index = 0;
        }

        public CellWriterHelper(MemberInfo member, int i)
        {
            Index = i;
            var memberType = member switch
            {
                PropertyInfo pi => pi.PropertyType,
                FieldInfo fi => fi.FieldType,
                _ => throw new InvalidOperationException()
            };
            Name = member switch
            {
                PropertyInfo pi => pi.Name,
                FieldInfo fi => fi.Name,
                _ => throw new InvalidOperationException()
            };
            Writer = FormatterHelperExtention.GenerateWriter<T>(typeof(T), memberType, Name);
        }
        public int Index { get; init; }
        public string Name { get; set; }
        public SerializeDelegate<T> Writer { get; init; }
    }

    public delegate int SerializeDelegate<T>(T value, ref IBufferWriter<byte> writer);
    public static class FormatterHelperExtention
    {
        readonly static Type _objectType = typeof(object);
        readonly static (MethodInfo method, Type? type)[] _methods =
            typeof(CellWriter)
                .GetMethods()
                .Where(m => m.Name == "Write")
                .Select(m => (m, m.GetParameters()?.FirstOrDefault()?.ParameterType))
                .ToArray();

        readonly static MethodInfo _methodObject = _methods.Where(x => x.type == typeof(object)).First().method;
        readonly static ParameterExpression _writer = Expression.Parameter(typeof(IBufferWriter<byte>).MakeByRefType(), "w");

        public static SerializeDelegate<T> GenerateWriter<T>(Type declaringType, Type propertyType, string name)
        {
            if (declaringType == null || propertyType.IsGenericType)
                return (T value, ref IBufferWriter<byte> writer) => CellWriter.WriteEmpty(ref writer);

            var methodTyped = _methods.Where(x => x.type == propertyType)?.Select(x => x.method)?.FirstOrDefault();
            var target = Expression.Parameter(declaringType, "i");
            var parameters = methodTyped == null
                ? new Expression[] { Expression.Convert(Expression.PropertyOrField(target, name), _objectType), _writer }
                : new Expression[] { Expression.PropertyOrField(target, name), _writer };

            var call = Expression.Call(methodTyped ?? _methodObject, parameters);
            var lambda = Expression.Lambda<SerializeDelegate<T>>(call, target, _writer);
            return lambda.Compile();
        }
    }
}
