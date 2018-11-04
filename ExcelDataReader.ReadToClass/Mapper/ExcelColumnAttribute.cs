using System;

namespace ExcelDataReader.ReadToClass.Mapper
{
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; }

        public int Order { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        /// <param name="order">The order.</param>
        public ExcelColumnAttribute(string name, int order)
        {
            Name = name;
            Order = order;
        }


        // ToDo: we need other way around:

        /// <summary>
        /// Gets the column letters from numeric index.
        /// E.g., 1=A, 1466 = BDJ.
        /// </summary>
        /// <param name="columnNr">The column nr (1-based).</param>
        private static string GetColumnLetters(int columnNr)
        {
            // 1 = A, 256 = IV, 419  = PC, 1466 = BDJ
            const int letterCount = 26;
            const int letterCount2 = letterCount * letterCount;
            const int baseLetter = 'A' - 1;

            var letter3 = columnNr / letterCount2;
            var letter3Rem = columnNr % letterCount2;

            var letter2 = letter3Rem / letterCount;
            var letter1 = letter3Rem % letterCount;

            var result = "";
            if (letter3 > 0)
                result += (char)(baseLetter + letter3);

            if (letter2 > 0)
                result += (char)(baseLetter + letter2);

            result += (char)(baseLetter + letter1);

            return result;
        }
    }
}
