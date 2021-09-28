using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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

            var beforeSlash1 = s1.Split('/').First();
            var beforeSlash2 = s2.Split('/').First();
            
            var isNumeric1 = IsNumeric(beforeSlash1);
            var isNumeric2 = IsNumeric(beforeSlash2);

            if (isNumeric1 && isNumeric2)
            {
                var i1 = Convert.ToInt32(beforeSlash1);
                var i2 = Convert.ToInt32(beforeSlash2);

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

            var sizes = new List<string> {"XXS", "XS", "XS/S", "S", "S/M", "M", "M/L", "L", "L/XL", "XL", "XXL", "XXXL"};

            var isInSizes1 = sizes.Contains(s1);
            var isInSizes2 = sizes.Contains(s2);

            if (isInSizes1 && isInSizes2)
            {
                var i1 = sizes.IndexOf(s1);
                var i2 = sizes.IndexOf(s2);

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
