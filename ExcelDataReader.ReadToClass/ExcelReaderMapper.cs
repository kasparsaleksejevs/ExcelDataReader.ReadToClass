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
        //// private readonly bool throwOnError = false; // ToDo: add better error handling

        delegate void PropertySetter(object instance, object propertyValue);

        delegate object ClassInstantiator();

        public T ReadAllWorksheets<T>(IExcelDataReader reader, FluentConfig config = null) where T : class, new()
        {
            var dataFromExcel = new T();
            var sheetProcessors = new List<SheetProcessingData>();

            List<TablePropertyData> tables = null;
            if (config != null)
                tables = config.Tables;
            else
                tables = GetTableProperties(typeof(T));

            foreach (var table in tables)
                sheetProcessors.Add(new SheetProcessingData(table, CreateSetPropertyAction(typeof(T), table.PropertyName))); //table.ExcelSheetName

            var hasNextResult = false;
            var readerWasReset = false;
            do
            {
                readerWasReset = false;
                var worksheetName = reader.Name;
                var currentSheetTables = sheetProcessors.Where(m => m.TablePropertyData.ExcelSheetName.Equals(worksheetName, StringComparison.OrdinalIgnoreCase) && m.Processed == false).ToList();
                if (currentSheetTables.Count > 0)
                {
                    var processor = currentSheetTables.First();
                    var result = ProcessTable(reader, processor.TablePropertyData);
                    processor.PropertySetter(dataFromExcel, result);
                    processor.Processed = true;

                    // handling cases (experimental), where multiple tables (classes) are bound to one excel sheet
                    if (currentSheetTables.Count > 1)
                    {
                        reader.Reset(); // ToDo: this will kill performance on large files. Reading should be refactored to be able to read one sheet to different tables simultaneously...
                        readerWasReset = true;
                    }
                }

                if (!readerWasReset)
                    hasNextResult = reader.NextResult();

            } while (hasNextResult);

            return dataFromExcel;
        }

        public List<string> Errors { get; } = new List<string>();

        private object ProcessTable(IExcelDataReader reader, TablePropertyData tablePropertyData, bool hasHeader = true)
        {
            if (!hasHeader)
                throw new Exception("Tables without headers are not yet supported!");

            var tableRowType = tablePropertyData.ListElementType;
            var tableRowImplementationType = tablePropertyData.ListElementTypeImplementation;

            var instanceCreator = CreateInstanceInitializationAction(tableRowType, tableRowImplementationType);
            var listAdder = CreateAddToListAction(tableRowType);

            var propertyProcessors = new Dictionary<string, PropertySetter>();
            foreach (var property in tablePropertyData.Columns)
                propertyProcessors.Add(property.ExcelColumnName, CreateSetPropertyAction(tableRowImplementationType ?? tableRowType, property.PropertyName));

            var columnOffset = 0;
            var rowOffset = 0;

            if (!string.IsNullOrEmpty(tablePropertyData.StartingCellAddress))
            {
                columnOffset = ColumnPropertyData.GetColumnIndexFromCellAddress(tablePropertyData.StartingCellAddress) - 1;
                rowOffset = ColumnPropertyData.GetRowIndexFromCellAddress(tablePropertyData.StartingCellAddress) - 1;
            }

            // reading header
            if (hasHeader)
            {
                for (int i = 0; i <= rowOffset; i++)
                    reader.Read();
            }

            var totalFieldCount = reader.FieldCount;
            var columns = new List<ExcelColumn>();
            for (int i = columnOffset; i < totalFieldCount; i++)
            {
                var columnName = reader.GetValue(i)?.ToString();
                var propertyData = tablePropertyData.Columns.FirstOrDefault(m => m.ExcelColumnName == columnName);

                if (!string.IsNullOrEmpty(columnName) && propertyData != null)
                    columns.Add(new ExcelColumn { ColumnIndex = i, ColumnName = columnName, IsMandatory = propertyData.IsMandatory, PropertyName = propertyData.PropertyName });

                if (columns.Count >= tablePropertyData.Columns.Count)
                    break;
            }

            if (tablePropertyData.OnHeaderRead != null)
            {
                var rowData = new object[totalFieldCount];
                for (int i = 0; i < totalFieldCount; i++)
                    rowData[i] = reader.GetValue(i);

                tablePropertyData.OnHeaderRead(rowData);
            }

            var fieldsToReadCount = columns.Count;

            object compatableList = Activator.CreateInstance(typeof(List<>).MakeGenericType(tableRowType));
            var rowIndex = rowOffset + 2; //we skipped header
            var readingShouldStop = false;

            while (reader.Read())
            {
                // create instance using lambdas:
                var instance = instanceCreator();

                foreach (var column in columns)
                {
                    var cellValue = reader.GetValue(column.ColumnIndex);
                    if (column.IsMandatory && (cellValue is null || cellValue.ToString() == string.Empty))
                    {
                        readingShouldStop = true;
                        break;
                    }

                    try
                    {
                        // set instance properties using lambdas
                        propertyProcessors[column.ColumnName](instance, cellValue);
                    }
                    catch (FormatException formatException)
                    {
                        propertyProcessors[column.ColumnName](instance, null);
                        this.Errors.Add($"{formatException.Message} - ['{column.ColumnName}':{rowIndex}]");
                    }
                    catch (NullReferenceException nullRefException)
                    {
                        throw new NullReferenceException($"{column.PropertyName} instance is null. Ensure that instance can be initialized (if it is an interface - specify implementation class with 'ImplementingClass()' method, if it is a nested property - ensure it is auto-initialized).");
                    }
                }

                if (readingShouldStop)
                    break;

                if (tablePropertyData.OnRowRead != null)
                {
                    var rowData = new object[totalFieldCount];
                    for (int i = 0; i < totalFieldCount; i++)
                        rowData[i] = reader.GetValue(i);

                    tablePropertyData.OnRowRead(instance, rowData);
                }

                listAdder(compatableList, instance);
                rowIndex++;
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
            var instanceParam = Expression.Parameter(typeof(object), "containerInstance");
            var valueParam = Expression.Parameter(typeof(object), "valueToSet");

            // instanceParam is 'object' and does not contain properties we need, so we convert it to our instance type (which it should be already)
            var instanceExpr = Expression.Convert(instanceParam, propertyContainerType);

            // getting reference to the static method that handles data type changes (double to decimal, etc)
            var changeTypeMethod = typeof(ExcelReaderMapper).GetMethod(nameof(ChangeToType), BindingFlags.NonPublic | BindingFlags.Static);

            if (propertyName.Contains("."))
                return CreateNestedSetPropertyAction(propertyContainerType, propertyName, instanceParam, valueParam, instanceExpr, changeTypeMethod);

            var propertyInfo = propertyContainerType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            var setMethodInfo = propertyInfo.GetSetMethod();
            var propertyType = propertyInfo.PropertyType;

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
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(nullableUnderlyingType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(instanceExpr, setMethodInfo, Expression.Convert(changeTypeExpr, propertyType));
                var emptyLambda = Expression.Empty();
                var checkExpr = Expression.IfThenElse(Expression.Equal(Expression.Constant(null, typeof(object)), valueParam), emptyLambda, setPropertyExpression);
                return Expression.Lambda<PropertySetter>(checkExpr, instanceParam, valueParam).Compile();
            }
            else
            {
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(propertyType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(instanceExpr, setMethodInfo, changeTypeExpr);
                return Expression.Lambda<PropertySetter>(setPropertyExpression, instanceParam, valueParam).Compile();
            }
        }

        /// <summary>
        /// Creates the 'set property' action for the specified class and property.
        /// </summary>
        /// <param name="propertyContainerType">Type of the class containing the property.</param>
        /// <param name="propertyName">Name of the property to set (e.g., 'SubClass.Property1').</param>
        /// <param name="instanceParam">The instance parameter.</param>
        /// <param name="valueParam">The value parameter.</param>
        /// <param name="instanceExpr">The instance expr.</param>
        /// <param name="changeTypeMethod">The change type method.</param>
        /// <returns>Compiled lambda expression to set property to specified value.</returns>
        private static PropertySetter CreateNestedSetPropertyAction(Type propertyContainerType, string propertyName, ParameterExpression instanceParam, ParameterExpression valueParam, UnaryExpression instanceExpr, MethodInfo changeTypeMethod)
        {
            var sourceType = propertyContainerType;
            var splitPropertyNames = propertyName.Split('.');
            PropertyInfo innerPropertyInfo = null;
            for (int i = 0; i < splitPropertyNames.Length; i++)
            {
                var innerPropertyName = splitPropertyNames[i];
                innerPropertyInfo = sourceType.GetProperty(innerPropertyName, BindingFlags.Instance | BindingFlags.Public);
                sourceType = innerPropertyInfo.PropertyType;
            }

            // at this point we have 'innerPropertyInfo' pointing to the target nested-property (e.g., 'Prop1' in 'mainClass.SubClass.Prop1')
            var propertyType = innerPropertyInfo.PropertyType;
            var setMethodInfo = innerPropertyInfo.GetSetMethod();

            // create expression from property names to reference target property container instance (e.g.,  'e => e.SubClass')
            var propertyAccessExpr = splitPropertyNames.Take(splitPropertyNames.Length - 1).Aggregate<string, Expression>(instanceExpr, (c, m) => Expression.Property(c, m));

            var nullableUnderlyingType = Nullable.GetUnderlyingType(propertyType);
            if (nullableUnderlyingType != null)
            {
                // IF (valueParam is null) THEN do_not_set_anything ELSE set_as_nullable_value
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(nullableUnderlyingType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(propertyAccessExpr, setMethodInfo, Expression.Convert(changeTypeExpr, propertyType));
                var emptyLambda = Expression.Empty();
                var checkExpr = Expression.IfThenElse(Expression.Equal(Expression.Constant(null, typeof(object)), valueParam), emptyLambda, setPropertyExpression);
                return Expression.Lambda<PropertySetter>(checkExpr, instanceParam, valueParam).Compile();
            }
            else
            {
                var changeTypeMethodGeneric = changeTypeMethod.MakeGenericMethod(innerPropertyInfo.PropertyType);
                var changeTypeExpr = Expression.Call(changeTypeMethodGeneric, valueParam);
                var setPropertyExpression = Expression.Call(propertyAccessExpr, setMethodInfo, changeTypeExpr);
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
        private static ClassInstantiator CreateInstanceInitializationAction(Type type, Type implementationType)
        {
            if (type.IsInterface && implementationType is null)
                throw new Exception($"Implementation type for the interface {type.Name} is not specified");

            ConstructorInfo ctor;
            if (type.IsInterface)
                ctor = implementationType.GetConstructors().FirstOrDefault();
            else
                ctor = type.GetConstructors().FirstOrDefault();

            if (ctor is null)
                throw new Exception($"There are no constructor for {implementationType?.Name ?? type.Name}");

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
                            ExcelSheetName = excelColumnAttribute.GetName(),
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

            public bool Processed { get; set; }
        }

        class ExcelColumn
        {
            public int ColumnIndex { get; set; }

            public string ColumnName { get; set; }

            public bool IsMandatory { get; set; }

            public string PropertyName { get; set; }
        }
    }
}
