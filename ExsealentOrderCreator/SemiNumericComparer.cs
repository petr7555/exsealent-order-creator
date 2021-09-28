using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExsealentOrderCreator
{
    public class SemiNumericComparer : IComparer<string>
    {
        /// <summary>
        /// Method to determine if a string is a number
        /// </summary>
        /// <param name="value">String to test</param>
        /// <returns>True if numeric</returns>
        private static bool IsNumeric(string value)
        {
            return int.TryParse(value, out _);
        }

        /// <inheritdoc />
        public int Compare(string s1, string s2)
        {
            const int s1GreaterThanS2 = 1;
            const int s2GreaterThanS1 = -1;

            var isNumeric1 = IsNumeric(s1);
            var isNumeric2 = IsNumeric(s2);

            if (isNumeric1 && isNumeric2)
            {
                var i1 = Convert.ToInt32(s1);
                var i2 = Convert.ToInt32(s2);

                return i1 - i2;
            }

            if (isNumeric1)
            {
                return s2GreaterThanS1;
            }

            if (isNumeric2)
            {
                return s1GreaterThanS2;
            }

            var sizes = new Dictionary<string, int>
            {
                {"XS", 0},
                {"S", 1},
                {"M", 2},
                {"L", 3},
                {"XL", 4},
                {"XXL", 5}
            };

            var isInSizes1 = sizes.ContainsKey(s1);
            var isInSizes2 = sizes.ContainsKey(s2);

            if (isInSizes1 && isInSizes2)
            {
                var i1 = sizes[s1];
                var i2 = sizes[s2];

                return i1 - i2;
            }

            if (isInSizes1)
            {
                return s2GreaterThanS1;
            }

            if (isInSizes2)
            {
                return s1GreaterThanS2;
            }
            
            return string.Compare(s1, s2, true, CultureInfo.InvariantCulture);
        }
    }
}
