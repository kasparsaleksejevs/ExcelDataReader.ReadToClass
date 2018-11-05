using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.Mapper
{
    internal class ExcelReaderMapper
    {
        public T ReadAllWorksheets<T>(IExcelDataReader reader) where T : class, new()
        {
            var dataFromExcel = new T();
            var sheetProcessors = new Dictionary<string, SheetProcessingData>();
            var tables = GetTableProperties(typeof(T));
            foreach (var item in tables)
            {
                var listElementType = item.ListElementType;
                sheetProcessors.Add(item.ExcelTableName, new SheetProcessingData(listElementType, CreateSetPropertyAction(typeof(T), item.PropertyName)));
            }

            do
            {
                var worksheetName = reader.Name;
                if (sheetProcessors.ContainsKey(worksheetName))
                {
                    var processor = sheetProcessors[worksheetName];
                    var result = ProcessTable(reader, processor.RowPropertyType);
                    processor.PropertySetter(dataFromExcel, result);
                }

            } while (reader.NextResult());

            return dataFromExcel;
        }

        delegate void PropertySetter(object instance, object propertyValue);
        delegate object ClassInstantiator();

        private object ProcessTable(IExcelDataReader reader, Type tableRowType, bool hasHeader = true)
        {
            var instanceCreator = CreateInstanceInitializationAction(tableRowType);
            var listAdder = CreateAddToListAction(tableRowType);

            // read all properties and compile set-value lambdas
            var properties = GetColumnProperties(tableRowType);

            ////var propertyProcessors = new Dictionary<string, Action<object, object>>(); // could improve things...
            var propertyProcessors = new Dictionary<string, PropertySetter>();
            foreach (var property in properties)
                propertyProcessors.Add(property.ExcelColumnName, CreateSetPropertyAction(tableRowType, property.PropertyName));

            // reading header
            if (hasHeader)
                reader.Read();

            var fieldCount = reader.FieldCount;
            var columns = new List<ExcelColumn>();
            for (int i = 0; i < fieldCount; i++)
            {
                var columnName = reader.GetString(i);
                if (!string.IsNullOrEmpty(columnName) && properties.Any(m => m.ExcelColumnName == columnName))
                    columns.Add(new ExcelColumn { ColumnIndex = i, ColumnName = columnName });
            }

            fieldCount = columns.Count;

            object compatableList = Activator.CreateInstance(typeof(List<>).MakeGenericType(tableRowType));
            while (reader.Read())
            {
                // create instance using lambdas:
                var instance = instanceCreator();

                for (int i = 0; i < columns.Count; i++)
                {
                    var column = columns[i];
                    var cellValue = reader.GetValue(column.ColumnIndex);

                    // set instance properties using lambdas
                    propertyProcessors[column.ColumnName](instance, cellValue);
                }

                listAdder(compatableList, instance);
            }

            return compatableList;
        }

        /// <summary>
        /// Helper function to change type of the cell value.
        /// Used in expressions, as it is PITA to do it purely in expressions.
        /// </summary>
        /// <typeparam name="TResult">The type of the result.</typeparam>
        /// <param name="value">The value.</param>
        /// <returns>Converted value to the specified type.</returns>
        private static TResult ChangeToType<TResult>(object value)
        {
            if (value is null)
                return default(TResult);

            var nullable = Nullable.GetUnderlyingType(typeof(TResult));
            if (nullable != null)
                return (TResult)Convert.ChangeType(value, nullable);

            return (TResult)Convert.ChangeType(value, typeof(TResult));
        }

        /// <summary>
        /// Creates the 'set property' action for the specified class and property.
        /// </summary>
        /// <param name="propertyContainerType">Type of the class containing the property.</param>
        /// <param name="propertyName">Name of the property to set.</param>
        /// <returns>Compiled lambda expression to set property to specified value.</returns>
        /// <remarks>https://stackoverflow.com/questions/3475199/how-to-dynamically-set-a-property-of-a-class-without-using-reflection-with-dyna</remarks>
        private static PropertySetter CreateSetPropertyAction(Type propertyContainerType, string propertyName)
        {
            var propertyInfo = propertyContainerType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            var setMethodInfo = propertyInfo.GetSetMethod();
            //var propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
            var propertyType = propertyInfo.PropertyType;

            var instanceParam = Expression.Parameter(typeof(object), "containerInstance");
            var valueParam = Expression.Parameter(typeof(object), "valueToSet");

            var instance = Expression.Convert(instanceParam, propertyContainerType);
            //Expression value;
            //if (propertyType.IsValueType)
            //    value = Expression.Convert(Expression.Unbox(valueParam, {typeof(valueParam)}), propertyType);  // <- obvoiusly this doesnt quite work as we need unbox from object to the boxed type. 
            //else
            //    value = Expression.Convert(valueParam, propertyType);
            //var callExpr = Expression.Call(instance, setMethodInfo, value);

            var changeTypeMethod = typeof(ExcelReaderMapper).GetMethod(nameof(ChangeToType), BindingFlags.NonPublic | BindingFlags.Static);
            var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(propertyType);
            var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
            var callExpr = Expression.Call(instance, setMethodInfo, changeTypeExpr);

            var lambda = Expression.Lambda<PropertySetter>(callExpr, instanceParam, valueParam);
            return lambda.Compile();
        }

        /// <summary>
        /// Creates the 'add to list' action for the specified list type.
        /// </summary>
        /// <param name="listElementType">Type of the list element.</param>
        /// <returns>Compiled lambda expression to add element to the list.</returns>
        private static PropertySetter CreateAddToListAction(Type listElementType)
        {
            var genericListType = typeof(List<>).MakeGenericType(listElementType);
            var addToListMethodInfo = genericListType.GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);

            var instanceParam = Expression.Parameter(typeof(object), "listInstance");
            var valueParam = Expression.Parameter(typeof(object), "listElementValue");

            var instance = Expression.Convert(instanceParam, genericListType);
            var value = Expression.Convert(valueParam, listElementType);
            var callExpr = Expression.Call(instance, addToListMethodInfo, value);

            var lambda = Expression.Lambda<PropertySetter>(callExpr, instanceParam, valueParam);
            return lambda.Compile();
        }

        /// <summary>
        /// Creates the instance initialization action. 
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>Compiled lambda.</returns>
        /// <remarks>https://rogerjohansson.blog/2008/02/28/linq-expressions-creating-objects/</remarks>
        private static ClassInstantiator CreateInstanceInitializationAction(Type type)
        {
            ConstructorInfo ctor = type.GetConstructors().First();

            NewExpression newExp = Expression.New(ctor);
            LambdaExpression lambda = Expression.Lambda<ClassInstantiator>(newExp);
            ClassInstantiator compiled = (ClassInstantiator)lambda.Compile();
            return compiled;
        }

        private static List<ColumnPropertyData> GetColumnProperties(Type type)
        {
            var lstProperties = new List<ColumnPropertyData>();
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var property in properties)
            {
                var excelColumnAttribute = property.GetCustomAttribute<ExcelColumnAttribute>(true);
                if (excelColumnAttribute != null)
                {
                    lstProperties.Add(new ColumnPropertyData
                    {
                        PropertyName = property.Name,
                        ExcelColumnName = excelColumnAttribute.Name,
                        Order = excelColumnAttribute.Order
                    });
                }
            }

            return lstProperties.OrderBy(m => m.Order).ToList();
        }

        private static List<TablePropertyData> GetTableProperties(Type type)
        {
            var lstProperties = new List<TablePropertyData>();
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var property in properties)
            {
                var propertyType = property.PropertyType;
                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    var listElementType = propertyType.GetGenericArguments()[0];

                    var excelColumnAttribute = property.GetCustomAttribute<ExcelTableAttribute>(true);
                    if (excelColumnAttribute != null)
                    {
                        lstProperties.Add(new TablePropertyData
                        {
                            PropertyName = property.Name,
                            ExcelTableName = excelColumnAttribute.GetName(),
                            ListElementType = listElementType
                        });
                    }
                }
            }

            return lstProperties.ToList();
        }

        class SheetProcessingData
        {
            public SheetProcessingData(Type rowPropertyType, PropertySetter propertySetter)
            {
                this.RowPropertyType = rowPropertyType;
                this.PropertySetter = propertySetter;
            }

            public Type RowPropertyType { get; set; }

            public PropertySetter PropertySetter { get; set; }
        }

        class ColumnPropertyData
        {
            public string PropertyName { get; set; }

            public string ExcelColumnName { get; set; }

            public int Order { get; set; }
        }

        class TablePropertyData
        {
            public string PropertyName { get; set; }

            public Type ListElementType { get; set; }

            public string ExcelTableName { get; set; }
        }

        class ExcelColumn
        {
            public int ColumnIndex { get; set; }

            public string ColumnName { get; set; }
        }
    }
}
