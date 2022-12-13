// See https://aka.ms/new-console-template for more information

using ExcelDataReader;
using ExcelExtractorDotNet;
using System.Text;
using System.Text.Json;

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var files = Directory.GetFiles(@"D:\Okul\Yüksek Lisans\Natural Language Processing\Homework\02.1 Knowledge Extraction Results\odev2_veriler\");


var verbExcels = new List<Excel>();
var nounExcels = new List<Excel>();

var nounExcelColumns = new List<ExcelColumn>
{
    new ExcelColumn { ColumnIndex = 1, ColumnName = "tanimi_nedir" },
    new ExcelColumn { ColumnIndex = 2, ColumnName = "nerede_bulunur" },
    new ExcelColumn { ColumnIndex = 3, ColumnName = "canli_cansiz" },
    new ExcelColumn { ColumnIndex = 4, ColumnName = "ust_kavrami_nedir" },
    new ExcelColumn { ColumnIndex = 5, ColumnName = "yaninda_neler_bulunur" },
    new ExcelColumn { ColumnIndex = 6, ColumnName = "icinde_neler_bulunur" },
    new ExcelColumn { ColumnIndex = 7, ColumnName = "hammaddesi_nedir" },
    new ExcelColumn { ColumnIndex = 8, ColumnName = "sekli_nasil" },
    new ExcelColumn { ColumnIndex = 9, ColumnName = "hacmi" },
    new ExcelColumn { ColumnIndex = 10, ColumnName = "agirlik" },
    new ExcelColumn { ColumnIndex = 11, ColumnName = "ne_ise_yarar" },
    new ExcelColumn { ColumnIndex = 12, ColumnName = "kim_kullanir" },
    new ExcelColumn { ColumnIndex = 13, ColumnName = "rengi" },
    new ExcelColumn { ColumnIndex = 14, ColumnName = "sifatlari" },
};

var verbExcelColumns = new List<ExcelColumn>
{
    new ExcelColumn { ColumnIndex = 1, ColumnName = "tanimi_nedir" },
    new ExcelColumn { ColumnIndex = 2, ColumnName = "kim_ne_yapar" },
    new ExcelColumn { ColumnIndex = 3, ColumnName = "kim_ne_ile_yapilir" },
    new ExcelColumn { ColumnIndex = 4, ColumnName = "nasil_yapilir" },
    new ExcelColumn { ColumnIndex = 5, ColumnName = "ney_kime_yapilir" },
    new ExcelColumn { ColumnIndex = 6, ColumnName = "ney_kimi_yapilir" },
    new ExcelColumn { ColumnIndex = 7, ColumnName = "nerede_yapilir" },
    new ExcelColumn { ColumnIndex = 8, ColumnName = "nicin_yapilir" },
    new ExcelColumn { ColumnIndex = 9, ColumnName = "ne_olunca_yapilir" },
    new ExcelColumn { ColumnIndex = 10, ColumnName = "yapinca_ne_olur" },
    new ExcelColumn { ColumnIndex = 11, ColumnName = "fiziksel_zihinsel" }
};

foreach (var file in files)
{
    var fileGuid = Guid.NewGuid().ToString();

    using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
    {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            reader.Read();

            do
            {
                var excel = new Excel
                {
                    BelongsTo = stream.Name.Substring(stream.Name.LastIndexOf("\\") + 1)
                };

                while (reader.Read())
                {
                    var allRowValues = reader.FieldCount;

                    if (allRowValues > 0)
                    {
                        var word = reader.GetString(0)?.Trim().ConvertToArffUsableFormat();

                        if (!string.IsNullOrEmpty(word))
                        {
                            var excelLine = new ExcelLine
                            {
                                Word = word
                            };

                            if (stream.Name.Contains("isim"))
                            {
                                foreach (var excelColumn in nounExcelColumns)
                                {
                                    var columnValue = reader.GetValue(excelColumn.ColumnIndex)?.ToString();

                                    if (!string.IsNullOrEmpty(columnValue))
                                    {
                                        excelLine.Columns.Add(new Column
                                        {
                                            Name = excelColumn.ColumnName,
                                            Values = columnValue.ReplaceTurkishCharactersWithEnglishEquivalents().Split(',').ToList()
                                        });
                                    }
                                }
                            }
                            else
                            {
                                foreach (var excelColumn in verbExcelColumns)
                                {
                                    var columnValue = reader.GetValue(excelColumn.ColumnIndex)?.ToString();

                                    if (!string.IsNullOrEmpty(columnValue))
                                    {
                                        excelLine.Columns.Add(new Column
                                        {
                                            Name = excelColumn.ColumnName,
                                            Values = columnValue.ReplaceTurkishCharactersWithEnglishEquivalents().Split(',').ToList()
                                        });
                                    }
                                }
                            }

                            if (excelLine.Columns.Count > 0)
                            {
                                excel.ExcelLines.Add(excelLine);
                            }
                        }
                    }
                }

                if (stream.Name.Contains("isim", StringComparison.OrdinalIgnoreCase))
                {
                    nounExcels.Add(excel);
                }
                else
                {
                    verbExcels.Add(excel);
                }
            }
            while (reader.NextResult());
        }
    }
}

var serialized = System.Text.Json.JsonSerializer.Serialize(verbExcels);

var verbDefinitionOnly = verbExcels
    .SelectMany(excel => excel.ExcelLines.Select(line => new ExcelLine { Word = line.Word, Columns = line.Columns.Where(col => col.Name == nounExcelColumns[0].ColumnName).ToList() }))
    .GroupBy(line => line.Word)
    .OrderBy(group => group.Key)
    .ToList();

var nounDefinitionOnly = nounExcels
    .SelectMany(excel => excel.ExcelLines.Select(line => new ExcelLine { Word = line.Word, Columns = line.Columns.Where(col => col.Name == nounExcelColumns[0].ColumnName).ToList() }))
    .GroupBy(line => line.Word)
    .OrderBy(group => group.Key)
    .ToList();

var verbExceptDefinition = verbExcels
    .SelectMany(excel => excel.ExcelLines.Select(line => new ExcelLine { Word = line.Word, Columns = line.Columns.Skip(1).Take(line.Columns.Count - 1).ToList() }))
    .GroupBy(line => line.Word)
    .OrderBy(group => group.Key)
    .ToList();

var nounExceptDefinition = nounExcels
    .SelectMany(excel => excel.ExcelLines.Select(line => new ExcelLine { Word = line.Word, Columns = line.Columns.Skip(1).Take(line.Columns.Count - 1).ToList() }))
    .GroupBy(line => line.Word)
    .OrderBy(group => group.Key)
    .ToList();

File.WriteAllText(@"D:\Okul\Yüksek Lisans\Natural Language Processing\Homework\02.1 Knowledge Extraction Results\results\verb_onlyDefinition.arff", ArffBuilder.BuildArff(verbDefinitionOnly, verbDefinitionOnly.SelectMany(val => val).SelectMany(line => line.Columns).DistinctBy(val => val.Name).ToList()));
File.WriteAllText(@"D:\Okul\Yüksek Lisans\Natural Language Processing\Homework\02.1 Knowledge Extraction Results\results\noun_onlyDefinition.arff", ArffBuilder.BuildArff(nounDefinitionOnly, nounDefinitionOnly.SelectMany(val => val).SelectMany(line => line.Columns).DistinctBy(val => val.Name).ToList()));
File.WriteAllText(@"D:\Okul\Yüksek Lisans\Natural Language Processing\Homework\02.1 Knowledge Extraction Results\results\verb_exceptDefinition.arff", ArffBuilder.BuildArff(verbExceptDefinition, verbExceptDefinition.SelectMany(val => val).SelectMany(line => line.Columns).DistinctBy(val => val.Name).ToList()));
File.WriteAllText(@"D:\Okul\Yüksek Lisans\Natural Language Processing\Homework\02.1 Knowledge Extraction Results\results\noun_exceptDefinition.arff", ArffBuilder.BuildArff(nounExceptDefinition, nounExceptDefinition.SelectMany(val => val).SelectMany(line => line.Columns).DistinctBy(val => val.Name).ToList()));