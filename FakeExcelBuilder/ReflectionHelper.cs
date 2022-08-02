using System.Reflection;

namespace FakeExcelBuilder
{
    public interface IAccessor
    {
        object? GetValue(object target);
        void SetValue(object target, object? value);
    }

    internal sealed class Accessor<TTarget, TProperty> : IAccessor
    {
        private readonly Func<TTarget, TProperty>? Getter;
        private readonly Action<TTarget, TProperty>? Setter;

        public Accessor(Func<TTarget, TProperty>? getter, Action<TTarget, TProperty>? setter)
        {
            Getter = getter;
            Setter = setter;
        }

        public object? GetValue(object target)
        {
            if (Getter == null)
                return null;
            return Getter((TTarget)target);
        }

        public void SetValue(object target, object? value)
        {
            if (Setter != null && value != null)
                Setter((TTarget)target, Accessor<TTarget, TProperty>.ChangeType(value));
        }

        private static TProperty ChangeType(object value)
        {
            var typeOfProperty = typeof(TProperty);

            if (typeOfProperty.IsGenericType && typeOfProperty.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                typeOfProperty = Nullable.GetUnderlyingType(typeOfProperty);
            }
            return (TProperty)Convert.ChangeType(value, typeOfProperty!);
        }
    }

    public static class DynamicPropertyAccessPropertyExtension
    {
        public static IAccessor GetAccessor(this PropertyInfo property)
        {
            Type getterDelegateType = typeof(Func<,>).MakeGenericType(property.DeclaringType!, property.PropertyType);
            var getMethod = property.GetGetMethod();
            Delegate? getter = getMethod switch
            {
                null => null,
                _ => Delegate.CreateDelegate(getterDelegateType, getMethod)
            };

            Type setterDelegateType = typeof(Action<,>).MakeGenericType(property.DeclaringType!, property.PropertyType);
            var setMethod = property.GetSetMethod();
            Delegate? setter = setMethod switch
            {
                null => null,
                _ => Delegate.CreateDelegate(setterDelegateType, setMethod)
            };

            Type accessorType = typeof(Accessor<,>).MakeGenericType(property.DeclaringType!, property.PropertyType);
            return (IAccessor)Activator.CreateInstance(accessorType, getter, setter)!;
        }
    }
}
