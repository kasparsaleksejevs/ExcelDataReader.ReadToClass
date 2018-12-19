using System;

namespace ExcelDataReader.ReadToClass.AttributeMapping
{
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }
    }
}
