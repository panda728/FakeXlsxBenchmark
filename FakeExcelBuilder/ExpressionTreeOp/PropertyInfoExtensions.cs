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
        public static bool IsSupportedType(Type t)
        {
            if (t.IsGenericType || t.IsArray)
                return false;
            if (t == typeof(string) || t == typeof(Guid) || t == typeof(Enum))
                return true;
            if (t == typeof(int) || t == typeof(long) || t == typeof(decimal) || t == typeof(double) || t == typeof(float))
                return true;
            if (t == typeof(DateTime) || t == typeof(DateOnly) || t == typeof(TimeOnly) || t == typeof(object))
                return true;
            return false;
        }

        public static Func<object, IBufferWriter<byte>, long>? GenerateEncodedGetterLambda<T>(this PropertyInfo propertyInfo)
        {
            // Build an equivalent process with ExpresstionTree.
            // Func<T, long> getCategoryId = (i,writer) => Formatter.GetBytes((i as T).CategoryId,writer);
            if (typeof(T) != propertyInfo.DeclaringType)
                return null;
            if (!IsSupportedType(propertyInfo.PropertyType))
                return null;

            var instanceObj = Expression.Parameter(typeof(object), "i"); 
            var instance = Expression.Convert(instanceObj, propertyInfo.DeclaringType);
            var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
            var property = Expression.Property(instance, propertyInfo);
            var ps = new Expression[] { property, writer };
            var method = typeof(ColumnFormatter).GetMethod("GetBytes", new Type[] { propertyInfo.PropertyType, typeof(IBufferWriter<byte>) });
            if (method == null)
                return null;

            var call = Expression.Call(method, ps);
            var d = Expression.Lambda(call, instanceObj, writer).Compile();
            return (Func<object, IBufferWriter<byte>, long>)d;
        }
        //public static Func<T, IBufferWriter<byte>, long>? GenerateEncodedGetterLambda2<T>(this PropertyInfo propertyInfo)
        //{
        //    // Build an equivalent process with ExpresstionTree.
        //    // Func<T, long> getCategoryId = (i,writer) => Formatter.GetBytes((i as T).CategoryId,writer);
        //    if (typeof(T) != propertyInfo.DeclaringType)
        //        return null;
        //    if (!IsSupportedType(propertyInfo.PropertyType))
        //        return null;

        //    var instance = Expression.Parameter(propertyInfo.DeclaringType, "i");
        //    var writer = Expression.Parameter(typeof(IBufferWriter<byte>), "writer");
        //    var property = Expression.Property(instance, propertyInfo);
        //    var ps = new Expression[] { property, writer };
        //    var method = typeof(ColumnFormatter).GetMethod("GetBytes", new Type[] { propertyInfo.PropertyType, typeof(IBufferWriter<byte>) });
        //    if (method == null)
        //        return null;

        //    var call = Expression.Call(method, ps);
        //    return (Func<T, IBufferWriter<byte>, long>)Expression.Lambda(call, instance, writer).Compile();
        //}

        //public static Func<object, object> GenerateGetterLambda<T>(this PropertyInfo propertyInfo)
        //{
        //    // Build an equivalent process with ExpresstionTree.
        //    //   Func<object, object> getCategoryId = (i) => (object)((i as T).CategoryId);
        //    if (typeof(T) != propertyInfo.DeclaringType)
        //        throw new ArgumentException();

        //    var objParameterExpr = Expression.Parameter(typeof(object), "i");
        //    var instanceExpr = Expression.TypeAs(objParameterExpr, propertyInfo.DeclaringType);
        //    var propertyExpr = Expression.Property(instanceExpr, propertyInfo);
        //    var propertyObjExpr = Expression.Convert(propertyExpr, typeof(object));
        //    return Expression.Lambda<Func<object, object>>(propertyObjExpr, objParameterExpr).Compile();
        //}
    }
}
