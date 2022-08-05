using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.PancakeInterop
{
    internal static class AssocAccessor
    {
        public static Func<object, List<object>> GetContentMethod{ get; private set; }
        public static Func<object, List<string>> GetNamesMethod{ get; private set; }
        public static Func<object> CtorMethod { get; private set; }

        public static Type AssocType { get; private set; }
        public static void BindToType(Type assocType)
        {
            AssocType = assocType;

            GetContentMethod = assocType
                .GetProperty("Value", BindingFlags.Public | BindingFlags.Instance)
                .GetterLambda<List<object>>();

            GetNamesMethod = assocType
                .GetProperty("Names", BindingFlags.NonPublic | BindingFlags.Instance)
                .GetterLambda<List<string>>();

            CtorMethod = assocType.GetConstructor(Type.EmptyTypes).CtorLambda();
        }

        private static Func<object, T> GetterLambda<T>(this PropertyInfo property)
        {
            var objParameterExpr = Expression.Parameter(typeof(object));
            var instanceExpr = Expression.TypeAs(objParameterExpr, property.DeclaringType);
            var valueExpr = Expression.Call(instanceExpr, property.GetMethod);
            var propertyObjExpr = Expression.Convert(valueExpr, typeof(T));
            return Expression.Lambda<Func<object, T>>(propertyObjExpr, objParameterExpr).Compile();
        }

        private static Func<object> CtorLambda(this ConstructorInfo method)
        {
            var valueExpr = Expression.New(method);
            var objExpr = Expression.Convert(valueExpr, typeof(object));
            return Expression.Lambda<Func<object>>(objExpr).Compile();
        }
    }
}