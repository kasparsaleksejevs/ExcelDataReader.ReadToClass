﻿using ExcelDataReader.ReadToClass.AttributeMapping;
using ExcelDataReader.ReadToClass.FluentMapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass
{
    internal class ExcelReaderMapper
    {
        private readonly bool throwOnError = false;

        public T ReadAllWorksheets<T>(IExcelDataReader reader, FluentConfig config = null) where T : class, new()
        {
            var dataFromExcel = new T();
            var sheetProcessors = new Dictionary<string, SheetProcessingData>();

            List<TablePropertyData> tables = null;
            if (config != null)
                tables = config.Tables;
            else
                tables = GetTableProperties(typeof(T));

            foreach (var table in tables)
                sheetProcessors.Add(table.ExcelTableName, new SheetProcessingData(table, CreateSetPropertyAction(typeof(T), table.PropertyName)));

            do
            {
                var worksheetName = reader.Name;
                if (sheetProcessors.ContainsKey(worksheetName))
                {
                    var processor = sheetProcessors[worksheetName];
                    var result = ProcessTable(reader, processor.TablePropertyData);
                    processor.PropertySetter(dataFromExcel, result);
                }

            } while (reader.NextResult());

            return dataFromExcel;
        }

        public List<string> Errors { get; } = new List<string>();

        delegate void PropertySetter(object instance, object propertyValue);
        delegate object ClassInstantiator();

        private object ProcessTable(IExcelDataReader reader, TablePropertyData tablePropertyData, bool hasHeader = true)
        {
            var tableRowType = tablePropertyData.ListElementType;

            var instanceCreator = CreateInstanceInitializationAction(tableRowType);
            var listAdder = CreateAddToListAction(tableRowType);

            var propertyProcessors = new Dictionary<string, PropertySetter>();
            foreach (var property in tablePropertyData.Columns)
                propertyProcessors.Add(property.ExcelColumnName, CreateSetPropertyAction(tableRowType, property.PropertyName));

            // reading header
            if (hasHeader)
                reader.Read();

            var fieldsToReadCount = reader.FieldCount;
            var totalFieldCount = reader.FieldCount;
            var columns = new List<ExcelColumn>();
            for (int i = 0; i < fieldsToReadCount; i++)
            {
                var columnName = reader.GetValue(i)?.ToString();
                if (!string.IsNullOrEmpty(columnName) && tablePropertyData.Columns.Any(m => m.ExcelColumnName == columnName))
                    columns.Add(new ExcelColumn { ColumnIndex = i, ColumnName = columnName });
            }

            if (tablePropertyData.onHeaderRead != null)
            {
                var rowData = new object[totalFieldCount];
                for (int i = 0; i < totalFieldCount; i++)
                    rowData[i] = reader.GetValue(i);

                tablePropertyData.onHeaderRead(rowData);
            }

            fieldsToReadCount = columns.Count;

            object compatableList = Activator.CreateInstance(typeof(List<>).MakeGenericType(tableRowType));
            var row = 2; //we skipped header
            while (reader.Read())
            {
                // create instance using lambdas:
                var instance = instanceCreator();

                for (int i = 0; i < columns.Count; i++)
                {
                    var column = columns[i];
                    var cellValue = reader.GetValue(column.ColumnIndex);

                    // set instance properties using lambdas
                    try
                    {
                        propertyProcessors[column.ColumnName](instance, cellValue);
                    }
                    catch (FormatException formatException)
                    {
                        propertyProcessors[column.ColumnName](instance, null);
                        this.Errors.Add($"{formatException.Message} - ['{column.ColumnName}':{row}]");
                    }
                }

                if (tablePropertyData.onRowRead != null)
                {
                    var rowData = new object[totalFieldCount];
                    for (int i = 0; i < totalFieldCount; i++)
                        rowData[i] = reader.GetValue(i);

                    tablePropertyData.onRowRead(instance, rowData);
                }

                listAdder(compatableList, instance);
                row++;
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
            {
                // nullable enums are handled in expressions, here we convert all other types
                return (TResult)Convert.ChangeType(value, nullable);
            }

            if (typeof(TResult) == typeof(DateTime) && value.GetType() == typeof(double))
                return (TResult)(object)DateTime.FromOADate((double)value);

            if (typeof(TResult).IsEnum)
            {
                if (value is string stringValue)
                    return (TResult)Enum.Parse(typeof(TResult), stringValue);

                return (TResult)Convert.ChangeType(value, typeof(int));
            }

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

            var nullableUnderlyingType = Nullable.GetUnderlyingType(propertyType);
            if (nullableUnderlyingType != null)
            {
                // IF (valueParam is null) THEN do_not_set_anything ELSE set_as_nullable_value
                var changeTypeMethod = typeof(ExcelReaderMapper).GetMethod(nameof(ChangeToType), BindingFlags.NonPublic | BindingFlags.Static);
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(nullableUnderlyingType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(instance, setMethodInfo, Expression.Convert(changeTypeExpr, propertyType));
                var emptyLambda = Expression.Empty();
                var checkExpr = Expression.IfThenElse(Expression.Equal(Expression.Constant(null, typeof(object)), valueParam), emptyLambda, setPropertyExpression);
                return Expression.Lambda<PropertySetter>(checkExpr, instanceParam, valueParam).Compile();
            }
            else
            {
                var changeTypeMethod = typeof(ExcelReaderMapper).GetMethod(nameof(ChangeToType), BindingFlags.NonPublic | BindingFlags.Static);
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(propertyType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(instance, setMethodInfo, changeTypeExpr);
                return Expression.Lambda<PropertySetter>(setPropertyExpression, instanceParam, valueParam).Compile();
            }
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
                    });
                }
            }

            return lstProperties;
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
                        var tableProperties = GetColumnProperties(listElementType);

                        lstProperties.Add(new TablePropertyData
                        {
                            PropertyName = property.Name,
                            ExcelTableName = excelColumnAttribute.GetName(),
                            ListElementType = listElementType,
                            Columns = tableProperties
                        });
                    }
                }
            }

            return lstProperties.ToList();
        }

        class SheetProcessingData
        {
            public SheetProcessingData(TablePropertyData tablePropertyData, PropertySetter propertySetter)
            {
                this.TablePropertyData = tablePropertyData;
                this.PropertySetter = propertySetter;
            }

            public TablePropertyData TablePropertyData { get; set; }

            public PropertySetter PropertySetter { get; set; }
        }

        class ExcelColumn
        {
            public int ColumnIndex { get; set; }

            public string ColumnName { get; set; }
        }
    }

    public class TablePropertyData
    {
        public string PropertyName { get; set; }

        public Type ListElementType { get; set; }

        public string ExcelTableName { get; set; }

        public List<ColumnPropertyData> Columns { get; set; } = new List<ColumnPropertyData>();

        public Action<object[]> onHeaderRead;

        public Action<object, object[]> onRowRead;
    }

    public class ColumnPropertyData
    {
        public string PropertyName { get; set; }

        public string ExcelColumnName { get; set; }
    }
}
