using ExcelDataReader.ReadToClass.FluentMapping;
using System.Collections.Generic;

namespace ExcelDataReader.ReadToClass
{
    /// <summary>
    /// ExcelDataReader extensions.
    /// </summary>
    /// <remarks>https://github.com/ExcelDataReader/ExcelDataReader</remarks>
    public static class ExcelDataReaderExtensions
    {
        /// <summary>
        /// Converts sheets to the specified class.
        /// </summary>
        /// <typeparam name="T">Type of class.</typeparam>
        /// <param name="reader">The IExcelDataReader instance.</param>
        /// <returns>Class {T} containing all data.</returns>
        /// <remarks>Note that sheet data properties musst be Lists (eg., List{someClass}), row types must have parameterless constructor.</remarks>
        public static T AsClass<T>(this IExcelDataReader reader, FluentConfig fluentConfig = null) where T : class, new()
        {
            return new ExcelReaderMapper().ReadAllWorksheets<T>(reader, fluentConfig);
        }

        /// <summary>
        /// Converts sheets to the specified class.
        /// </summary>
        /// <typeparam name="T">Type of class.</typeparam>
        /// <param name="reader">The IExcelDataReader instance.</param>
        /// <param name="errors">The errors that were encountered during read (invalid format, etc).</param>
        /// <returns>Class {T} containing all data.</returns>
        /// <remarks>Note that sheet data properties musst be Lists (eg., List{someClass}), row types must have parameterless constructor.</remarks>
        public static T AsClass<T>(this IExcelDataReader reader, out List<string> errors, FluentConfig fluentConfig = null) where T : class, new()
        {
            var excelMapper = new ExcelReaderMapper();
            var data = excelMapper.ReadAllWorksheets<T>(reader, fluentConfig);
            errors = excelMapper.Errors;
            return data;
        }
    }
}
