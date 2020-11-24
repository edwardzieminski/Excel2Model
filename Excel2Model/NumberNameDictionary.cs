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
    }
}
