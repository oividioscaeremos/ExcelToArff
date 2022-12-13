using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractorDotNet
{
    internal class ArffBuilder
    {
        public enum ArffType
        {
            DefinitionOnly,
            ExceptDefinition,
            All
        }

        public static string BuildArff(List<IGrouping<string, ExcelLine>> values, List<Column> columns)
        {
            var sb = new StringBuilder("@RELATION\ttext2arff\n\n");

            foreach (var column in columns)
            {
                sb.AppendLine("@ATTRIBUTE\t" + column.Name + "\tSTRING");
            }

            sb.AppendLine($"@ATTRIBUTE\tclass\t{{{string.Join(",", values.SelectMany(v => v.Select(val => val.Word.Replace(' ', '_'))).DistinctBy(val => val)).ReplaceTurkishCharactersWithEnglishEquivalents()}}}");
            sb.AppendLine();
            sb.AppendLine("@DATA");

            foreach (var group in values)
            {
                foreach (var line in group)
                {
                    foreach (var column in columns)
                    {
                        var vals = line.Columns.FirstOrDefault(col => col.Name == column.Name)?.Values;

                        var value = vals != null ? string.Join(" ", vals) : string.Empty;

                        sb.Append("\"" + value?.ReplaceTurkishCharactersWithEnglishEquivalents() + "\"" + ",");
                    }

                    sb.Append(line.Word);
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }
    }
}