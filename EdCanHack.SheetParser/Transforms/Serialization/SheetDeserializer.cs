using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EdCanHack.SheetParser.Transforms.Serialization
{
    public class SheetDeserializer<TDeserializationType> : SheetTransformer<TDeserializationType>
        where TDeserializationType : class, new()
    {
        public readonly Type DeserializationType;
        public readonly TypeInfo DeserializationTypeInfo;

        private readonly Dictionary<String, PropertyInfo> _propertyMappings;

        private readonly Boolean _ignoreMissingFields;
        private readonly Boolean _ignoreExtraFields;

        private readonly Dictionary<Type, Func<String, Object>> _parsers;

        public SheetDeserializer(IDictionary<Type, Func<String, Object>> extraParsers = null)
        {
            DeserializationType = typeof (TDeserializationType);
            DeserializationTypeInfo = typeof (TDeserializationType).GetTypeInfo();

            if (!CanDeserializeType(DeserializationType))
            {
                throw new SheetDeserializationException("The requested class can't be deserialized with SheetParser.");
            }

            _ignoreMissingFields = ShouldIgnoreMissingFields(DeserializationTypeInfo);
            _ignoreExtraFields = ShouldIgnoreExtraFields(DeserializationTypeInfo);

            if (extraParsers == null)
            {
                _parsers = DefaultParsers;
            }
            else
            {
                _parsers = new Dictionary<Type, Func<string, object>>(extraParsers);
                foreach (var kvp in extraParsers.Where(kvp => !_parsers.ContainsKey(kvp.Key)))
                    _parsers.Add(kvp.Key, kvp.Value);
            }

            _propertyMappings = BuildMapping();
        }


        public static Boolean CanDeserializeType<TTestType>()
        {
            return CanDeserializeType(typeof (TTestType));
        }

        public static Boolean CanDeserializeType(Type type)
        {
            // TODO: allow DataContract or automatic discovery (But I hate automatic discovery, so...patches welcome.)
            var info = type.GetTypeInfo();

            return info.GetCustomAttribute<SheetDeserializableAttribute>() != null;
        }

        private static Boolean ShouldIgnoreMissingFields(TypeInfo info)
        {
            var attr = info.GetCustomAttribute<SheetDeserializableAttribute>();
            return attr != null && attr.IgnoreMissingFields;
        }

        private static Boolean ShouldIgnoreExtraFields(TypeInfo info)
        {
            var attr = info.GetCustomAttribute<SheetDeserializableAttribute>();
            return attr == null || attr.IgnoreExtraFields;
        }


        private Dictionary<String, PropertyInfo> BuildMapping()
        {
            var propertyMapping = new Dictionary<String, PropertyInfo>();
            Boolean isKeyed = DeserializationTypeInfo.ImplementedInterfaces.Any(iface =>
            {
                var ifaceInfo = iface.GetTypeInfo();

                return ifaceInfo.IsGenericType && ifaceInfo.GetGenericTypeDefinition() == typeof (IKeyed<>);
            });

            var enumerable = DeserializationTypeInfo.DeclaredProperties;
            if (isKeyed)
            {
                var keyProp = EnsureValidSerializationProperty(enumerable.First(prop => prop.Name == "Key"));


                propertyMapping.Add("Key", keyProp);
                enumerable = enumerable.Where(prop => prop.Name != "Key");
            }


            foreach (var prop in enumerable)
            {
                var attr = prop.GetCustomAttribute<SheetPropertyAttribute>();
                if (attr == null) continue;

                propertyMapping.Add(attr.ColumnName, EnsureValidSerializationProperty(prop));
            }

            return propertyMapping;
        }


        public override IEnumerable<TDeserializationType> TransformRows(Sheet sheet)
        {
            Dictionary<String, Int32> headerMappings = FindHeaderMappings(sheet);

            var sheetRows = sheet.EnumerateRows();
            var transformRows = sheetRows.Select(r => FromRowImpl(headerMappings, r));
            return transformRows;
        }

        protected override TDeserializationType FromRow(IList<string> row)
        {
            throw new NotImplementedException(
                "The deserializer doesn't use FromRow directly because it needs another closure.");
        }


        private static readonly Object[] EmptyObjectArray = {};
        protected TDeserializationType
            FromRowImpl(IDictionary<String, Int32> headerMappings, IList<String> row)
        {
#if !DEBUG
            try
            {
#endif
                var retval =
                    DeserializationTypeInfo.DeclaredConstructors.First(ctor => !ctor.GetParameters().Any())
                        .Invoke(EmptyObjectArray);

                foreach (var propertyMapping in _propertyMappings)
                {
                    String propertyName = propertyMapping.Key;
                    Int32 column;
                    if (!headerMappings.TryGetValue(propertyName, out column)) continue;

                    String cell = row[column];

                    if (String.IsNullOrWhiteSpace(cell)) continue;
                    
                    var value = _parsers[propertyMapping.Value.PropertyType].Invoke(cell);
                    propertyMapping.Value.SetValue(retval, value);
                }

                return (TDeserializationType) retval;
#if !DEBUG
            }
            catch (Exception ex)
            {
                throw new SheetDeserializationException("Error when deserializing row: {0}",
                    String.Join(", ", row.Select(s => String.Format("'{0}'", s))));
            }
#endif
        }

        private Dictionary<String, Int32> FindHeaderMappings(Sheet sheet)
        {
            var header = sheet.HeaderRow;
            if (header == null) throw new SheetDeserializationException("Cannot deserialize from a headerless sheet.");

            var mappings = new Dictionary<string, int>(header.Count);
            for (var i = 0; i < header.Count; ++i)
            {
                var name = header[i];
                if (String.IsNullOrWhiteSpace(name)) continue;

                mappings.Add(name, i);
            }

            if (!_ignoreMissingFields)
            {
                var missing = _propertyMappings.Where(m => !mappings.ContainsKey(m.Key)).ToList();
                if (missing.Count > 0)
                {
                    throw new SheetDeserializationException("Missing columns not found in sheet: {0}",
                        String.Join(", ", missing));
                }
            }
            if (!_ignoreExtraFields)
            {
                var extra = mappings.Where(m => !_propertyMappings.ContainsKey(m.Key)).ToList();
                if (extra.Count > 0)
                {
                    throw new SheetDeserializationException("Extra columns found in sheet: {0}",
                        String.Join(", ", extra));
                }
            }

            return mappings;
        }

        private PropertyInfo EnsureValidSerializationProperty(PropertyInfo prop)
        {
            if (!prop.CanRead || !prop.CanWrite)
                throw new SheetDeserializationException("Property '{0}' must be readable and writable.", prop.Name);

            if (!_parsers.ContainsKey(prop.PropertyType))
                throw new SheetDeserializationException(
                    "Property '{0}' is of type '{1}', which is not in the parser list.", prop.Name,
                    prop.PropertyType.FullName);

            return prop;
        }


        private static readonly Dictionary<Type, Func<String, Object>> DefaultParsers = new Dictionary
            <Type, Func<String, Object>>()
        {
            {typeof (Int16), s => Int16.Parse(s)},
            {typeof (UInt16), s => UInt16.Parse(s)},
            {typeof (Int32), s => Int32.Parse(s)},
            {typeof (UInt32), s => UInt32.Parse(s)},
            {typeof (Int64), s => Int64.Parse(s)},
            {typeof (UInt64), s => UInt64.Parse(s)},
            {typeof (Decimal), s => Decimal.Parse(s)},
            {typeof (Single), s => Single.Parse(s)},
            {typeof (Double), s => Double.Parse(s)},
            {typeof (String), s => s.Trim() }

        };
    }
}
