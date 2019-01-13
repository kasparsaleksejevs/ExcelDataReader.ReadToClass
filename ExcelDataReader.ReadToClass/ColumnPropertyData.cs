using System;

namespace ExcelDataReader.ReadToClass
{
    public class ColumnPropertyData
    {
        public string PropertyName { get; set; }

        public string ExcelColumnName { get; set; }

        public bool IsMandatory { get; set; }

        /// <summary>
        /// Gets the column index (1-based) from column address (e.g., A=1, BDJ=1466).
        /// Can be specified with or without row index (both 'AB12' and 'AB' are valid).
        /// </summary>
        /// <param name="cellAddress">The cell address.</param>
        /// <returns>Column index (1-based).</returns>
        public static int GetColumnIndexFromCellAddress(string cellAddress)
        {
            // A=1, IV=256, PC=419, BDJ=1466 
            cellAddress = cellAddress.ToUpper();
            const int letterA = 'A' - 1;
            const int letterZ = 'Z';
            const int letterBase = 26;

            var addressLength = 0;
            for (int i = 0; i < cellAddress.Length; i++)
            {
                if (cellAddress[i] <= letterA || cellAddress[i] > letterZ)
                    break;

                addressLength++;
            }

            cellAddress = cellAddress.Substring(0, addressLength);

            var result = 0;
            for (int i = 0; i < cellAddress.Length; i++)
            {
                var power = 1;
                for (int j = 0; j < cellAddress.Length - i - 1; j++)
                    power *= letterBase;

                result += (cellAddress[i] - letterA) * power;
            }

            return result;
        }

        /// <summary>
        /// Gets the row index (1-based) from cell address.
        /// The address must be specified with both column letters and row index.
        /// </summary>
        /// <param name="cellAddress">The cell address.</param>
        /// <returns>Column index (1-based).</returns>
        public static int GetRowIndexFromCellAddress(string cellAddress)
        {
            // A=1, IV=256, PC=419, BDJ=1466 
            cellAddress = cellAddress.ToUpper();
            const int letterA = 'A';
            const int letterZ = 'Z';

            var resultString = string.Empty;
            for (int i = 0; i < cellAddress.Length; i++)
            {
                if (cellAddress[i] >= letterA && cellAddress[i] <= letterZ)
                    continue;

                resultString += cellAddress[i];
            }

            return Convert.ToInt32(resultString);
        }
    }
}
