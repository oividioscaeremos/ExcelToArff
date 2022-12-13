using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractorDotNet
{
    public static class ExtensionMethods
    {
        public static string ReplaceTurkishCharactersWithEnglishEquivalents(this string val)
        {
            if (string.IsNullOrEmpty(val))
                return string.Empty;

            return val
                .Replace('ç', 'c')
                .Replace('ğ', 'g')
                .Replace('ı', 'i')
                .Replace('ö', 'o')
                .Replace('ş', 's')
                .Replace('ü', 'u')
                .Replace('Ç', 'c')
                .Replace('Ğ', 'g')
                .Replace('Ö', 'o')
                .Replace('Ş', 's')
                .Replace('Ü', 'u')
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
        }

        public static string ConvertToArffUsableFormat(this string val)
        {
            return val.Trim().ReplaceTurkishCharactersWithEnglishEquivalents()
                .Replace(" ", "_")
                .Replace("?", "")
                .Replace("/", "");
        }
    }


}
