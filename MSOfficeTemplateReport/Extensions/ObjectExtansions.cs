using System;
using System.Collections.Generic;
using System.Reflection;
using System.Reflection.Emit;
using System.Text.Json;

namespace MSOfficeTemplateReport.Extensions
{
    public static class ObjectExtensions
    {
        public static object JsonElementToObjectObject(this string json)
        {
            var element = JsonDocument.Parse(json).RootElement;

            if (element.ValueKind == JsonValueKind.Array)
            {
                List<object> result = new List<object>();
                foreach (var s in element.EnumerateArray())
                {
                    var o = s.ToString().JsonElementToObjectObject();
                    result.Add(o);
                }
                return result.ToArray();
            }

            var schema = InferSchema(element);

            var type = GenerateType("DynamicJsonType", schema);

            var instance = CreateInstance(type, element);

            return instance;
        }

        static Dictionary<string, Type> InferSchema(JsonElement element)
        {
            var dict = new Dictionary<string, Type>();

            foreach (var prop in element.EnumerateObject())
            {
                Type type;
                switch (prop.Value.ValueKind)
                {                    
                    case JsonValueKind.Object:
                        type = typeof(object);
                        break;
                    case JsonValueKind.Array:
                        type = typeof(List<object>);
                        break;
                    case JsonValueKind.String:
                        type = typeof(string);
                        break;
                    case JsonValueKind.Number:
                        type = prop.Value.TryGetInt64(out _) ? typeof(long) : typeof(double);
                        break;
                    case JsonValueKind.True:
                        type = typeof(bool);
                        break;
                    case JsonValueKind.False:
                        type = typeof(string);
                        break;                   
                    default:
                        type = typeof(object);
                        break;
                }               
                dict[prop.Name] = type;
            }
            return dict;
        }

        static Type GenerateType(string typeName, Dictionary<string, Type> props)
        {
            var asmName = new AssemblyName("DynamicJsonAssembly");
            var asm = AssemblyBuilder.DefineDynamicAssembly(asmName, AssemblyBuilderAccess.Run);
            var module = asm.DefineDynamicModule("MainModule");

            var tb = module.DefineType(typeName, TypeAttributes.Public);

            foreach (var item in props)
            {

                var field = tb.DefineField($"_{item.Key}", item.Value, FieldAttributes.Private);

                var propBuilder = tb.DefineProperty(item.Key, PropertyAttributes.HasDefault, item.Value, null);

                var get = tb.DefineMethod($"get_{item.Key}",
                    MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig,
                    item.Value, Type.EmptyTypes);

                var ilGet = get.GetILGenerator();
                ilGet.Emit(OpCodes.Ldarg_0);
                ilGet.Emit(OpCodes.Ldfld, field);
                ilGet.Emit(OpCodes.Ret);

                var set = tb.DefineMethod($"set_{item.Key}",
                    MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig,
                    null, new Type[] { item.Value });

                var ilSet = set.GetILGenerator();
                ilSet.Emit(OpCodes.Ldarg_0);
                ilSet.Emit(OpCodes.Ldarg_1);
                ilSet.Emit(OpCodes.Stfld, field);
                ilSet.Emit(OpCodes.Ret);

                propBuilder.SetGetMethod(get);
                propBuilder.SetSetMethod(set);
            }

            return tb.CreateTypeInfo();
        }

        static object CreateInstance(Type type, JsonElement element)
        {
            var obj = Activator.CreateInstance(type);

            foreach (var prop in element.EnumerateObject())
            {
                var p = type.GetProperty(prop.Name);
                if (p == null) continue;

                object value;

                switch (prop.Value.ValueKind)
                {              
                    case JsonValueKind.String:
                        value = prop.Value.GetString();
                        break;
                    case JsonValueKind.Number:
                        value = GetNumberValue(prop.Value, p.PropertyType);
                        break;
                    case JsonValueKind.True:
                        value = true;
                        break;
                    case JsonValueKind.False:
                        value = false;
                        break;
                    default:
                        value = null;
                        break;
                }
                             
                p.SetValue(obj, value);
            }
            return obj;
        }

        static object GetNumberValue(JsonElement element, Type targetType)
        {
            if (targetType == typeof(long))
                return element.GetInt64();
            if (targetType == typeof(int))
                return element.GetInt32();
            if (targetType == typeof(float))
                return element.GetSingle();
            if (targetType == typeof(double))
                return element.GetDouble();
            if (targetType == typeof(decimal))
                return element.GetDecimal();

            return element.GetDouble();
        }
    }
}
