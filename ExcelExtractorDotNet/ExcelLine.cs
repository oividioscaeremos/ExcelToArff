using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractorDotNet
{
    internal class Excel
    {
        public string BelongsTo { get; set; }
        public List<ExcelLine> ExcelLines { get; set; }

        public Excel()
        {
            this.ExcelLines = new();
        }
    }

    internal class ExcelLine
    {
        public string Word { get; set; }
        public List<Column> Columns { get; set; }

        public ExcelLine()
        {
            this.Columns = new();
        }
    }

    internal class Column
    {
        public string Name { get; set; }
        public List<string> Values { get; set; }
    }

    internal class WordDefinition
    {
        public string Word { get; set; }
        public List<Column> Columns { get; set; }
    }

    internal class ExcelColumn
    {
        public string ColumnName { get; set; }
        public int ColumnIndex { get; set; }
    }
}
