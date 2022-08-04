using System;
using System.Buffers;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace FakeExcelBuilder.ExpressionTreeOp
{
    public static class PropertyInfoExtensions
    {
        public static Func<T, IBufferWriter<byte>, long>? GenerateEncodedGetterLambda<T>(this PropertyInfo propertyInfo)
        {
            // Build an equivalent process with ExpresstionTree.
            // Func<T, long> getCategoryId = (i,writer) => Formatter.GetBytes((i as T).CategoryId,writer);
            if (typeof(T) != propertyInfo.DeclaringType)
                throw new ArgumentException(nameof(propertyInfo));

            var instance = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var property = Expression.Property(instance, propertyInfo);
            var ps = new Expression[] { property, writer };
            var method = typeof(ColumnFormatter<T>).GetMethod("GetBytes", new Type[] { propertyInfo.PropertyType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return null;

            var callExpr = Expression.Call(method, ps);
            return (Func<T, IBufferWriter<byte>, long>)Expression.Lambda(callExpr, instance, writer).Compile();
        }

        public static Func<object, object> GenerateGetterLambda<T>(this PropertyInfo propertyInfo)
        {
            // Build an equivalent process with ExpresstionTree.
            //   Func<T, object> getCategoryId = (i) => (object)((i as T).CategoryId);
            if (typeof(T) != propertyInfo.DeclaringType)
                throw new ArgumentException();

            //var instance = Expression.Parameter(propertyInfo.DeclaringType, "i");
            //var property = Expression.Property(instance, propertyInfo);
            //var convert = Expression.TypeAs(property, typeof(object));
            //return (Func<object, object>)Expression.Lambda(convert, instance).Compile();


            var objParameterExpr = Expression.Parameter(typeof(object), "i");
            var instanceExpr = Expression.TypeAs(objParameterExpr, propertyInfo.DeclaringType);
            var propertyExpr = Expression.Property(instanceExpr, propertyInfo);
            var propertyObjExpr = Expression.Convert(propertyExpr, typeof(object));
            return Expression.Lambda<Func<object, object>>(propertyObjExpr, objParameterExpr).Compile();
        }
    }
}
