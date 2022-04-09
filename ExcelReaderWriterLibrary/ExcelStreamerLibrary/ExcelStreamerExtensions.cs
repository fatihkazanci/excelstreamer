using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary
{
    internal class ExcelStreamerExtensions
    {
        internal static IEnumerable<string> Generate()
        {
            long n = 0;
            while (true) yield return ToBase26(++n);
        }
        internal static string ToBase26(long i)
        {
            if (i == 0) return ""; i--;
            return ToBase26(i / 26) + (char)('A' + i % 26);
        }
        internal static string GetAString(string aa)
        {
            var news = aa.ToArray();
            string newItem = string.Empty;
            foreach (var item in news)
            {
                bool isAlphaBet = Regex.IsMatch(item.ToString(), "[a-z]", RegexOptions.IgnoreCase);
                if (isAlphaBet)
                {
                    newItem += item;
                }
                else
                {
                    break;
                }
            }
            return newItem;
        }
        internal static string GetNString(string aa)
        {
            var news = aa.ToArray();
            string newItem = string.Empty;
            foreach (var item in news)
            {
                bool isNumber = Regex.IsMatch(item.ToString(), "^[0-9]*$");
                if (isNumber)
                {
                    newItem += item;
                }
            }
            return newItem;
        }
    }
}
