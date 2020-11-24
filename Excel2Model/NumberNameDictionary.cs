using System;
using System.Collections.Generic;
using System.Linq;

namespace Excel2Model
{
    public static class NumberNameDictionary
    {
        private static readonly Dictionary<int, char> _nameNumberPairs;

        static NumberNameDictionary()
        {
            _nameNumberPairs = new Dictionary<int, char>()
            {
                { 0, default },
                { 1, 'A' },
                { 2, 'B' },
                { 3, 'C' },
                { 4, 'D' },
                { 5, 'E' },
                { 6, 'F' },
                { 7, 'G' },
                { 8, 'H' },
                { 9, 'I' },
                { 10,'J' },
                { 11, 'K' },
                { 12, 'L' },
                { 13, 'M' },
                { 14, 'N' },
                { 15, 'O' },
                { 16, 'P' },
                { 17, 'Q' },
                { 18, 'R' },
                { 19, 'S' },
                { 20, 'T' },
                { 21, 'U' },
                { 22, 'V' },
                { 23, 'W' },
                { 24, 'X' },
                { 25, 'Y' },
                { 26, 'Z' }
            };
        }

        public static int GetIndexOfLetter(char letter)
        {
            return _nameNumberPairs
                .Where(x => x.Value == letter)
                .Select(x => x.Key)
                .FirstOrDefault();
        }

        public static char GetLetterByIndex(int index)
        {
            return _nameNumberPairs
                .Where(x => x.Key == index)
                .Select(x => x.Value)
                .FirstOrDefault();
        }

        /// <summary>
        /// Supports column names of 1 or 2 capital letters.
        /// </summary>
        public static int GetColumnIndexByColumnName(string columnName)
        {
            int firstLetterIndex;
            int secondLetterIndex;

            if (columnName.Length == 1)
            {
                firstLetterIndex = GetIndexOfLetter(default);
                secondLetterIndex = GetIndexOfLetter(columnName.ToUpper()[0]);
            }
            else if (columnName.Length == 2)
            {
                firstLetterIndex = GetIndexOfLetter(columnName.ToUpper()[0]);
                secondLetterIndex = GetIndexOfLetter(columnName.ToUpper()[1]);
            }
            else
            {
                throw new Exception("Incorrect columName provided.");
            }

            return (firstLetterIndex * 26) + secondLetterIndex;
        }

        public static string GetColumnNameByColumnIndex(int columnIndex)
        {
            char firstLetter = GetLetterByIndex((columnIndex-1) / 26);
            char secondLetter = GetLetterByIndex(columnIndex % 26);

            return $"{firstLetter}{secondLetter}";
        }
    }
}
