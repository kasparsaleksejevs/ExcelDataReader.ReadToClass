using System;

namespace ExcelDataReader.ReadToClass.Mapper
{
    internal class ExcelTableAttribute : Attribute
    {
        private string name = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelTableAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the worksheet.</param>
        public ExcelTableAttribute(string name)
        {
            this.name = name;
        }

        public string GetName()
        {
            return this.name;
        }
    }
}
