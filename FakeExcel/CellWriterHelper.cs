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
            Write = (ref IBufferWriter<byte> writer, T value, CellWriter cellWriter) => cellWriter.Write(value, ref writer);
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
            Write = FormatterHelperExtention.CompiledWriter<T>(typeof(T), memberType, Name);
        }
        public int Index { get; init; }
        public string Name { get; set; }
        public SerializeDelegate<T> Write { get; init; }
    }

    public delegate int SerializeDelegate<T>(ref IBufferWriter<byte> writer, T value, CellWriter cellWriter);
    public static class FormatterHelperExtention
    {
        readonly static Type _objectType = typeof(object);
        readonly static (MethodInfo method, Type? type)[] _methods =
            typeof(CellWriter)
                .GetMethods()
                .Where(m => m.Name == "Write")
                .Select(m => (m, m.GetParameters()?.FirstOrDefault()?.ParameterType))
                .ToArray();

        readonly static MethodInfo _methodObjectWriter = _methods.Where(x => x.type == typeof(object)).First().method;
        readonly static ParameterExpression _writer = Expression.Parameter(typeof(IBufferWriter<byte>).MakeByRefType(), "w");
        readonly static ParameterExpression _cellWriter = Expression.Parameter(typeof(CellWriter), "c");

        public static SerializeDelegate<T> CompiledWriter<T>(Type declaringType, Type propertyType, string name)
        {
            // (i, ref w, c) => c.Write(i.Id, ref w)
            // (i, ref w, c) => c.Write((object)i.Id, ref w)
            if (propertyType.IsGenericType)
                return (ref IBufferWriter<byte> writer, T value, CellWriter cellWriter) => cellWriter.WriteEmpty(ref writer);

            var target = Expression.Parameter(declaringType, "i");

            var methodTyped = _methods.Where(x => x.type == propertyType)?.Select(x => x.method)?.FirstOrDefault();
            var parameters = methodTyped == null
                ? new Expression[] { Expression.Convert(Expression.PropertyOrField(target, name), _objectType), _writer }
                : new Expression[] { Expression.PropertyOrField(target, name), _writer };

            var call = Expression.Call(_cellWriter, methodTyped ?? _methodObjectWriter, parameters);
            var lambda = Expression.Lambda<SerializeDelegate<T>>(call, _writer, target, _cellWriter);
            return lambda.Compile();
        }
    }
}
