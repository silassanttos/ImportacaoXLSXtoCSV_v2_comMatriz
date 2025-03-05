using ClosedXML.Excel;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ImportacaoXLSXtoCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine(@"Uso correto: Exec xp_cmdshell ImportaXLSXtoCSV.exe ""<caminho_XLSX>"" ""<caminho_CSV>""");
                return;
            }

            string xlsxPath = args[0];
            string csvPath = args[1];

            try
            {
                Console.WriteLine($"Convertendo {xlsxPath} para {csvPath}...");
                ConvertXlsxToCsv(xlsxPath, csvPath);
                Console.WriteLine("Conversão concluída com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro durante a conversão: {ex.Message}");
            }
        }

        private static void ConvertXlsxToCsv(string xlsxPath, string csvPath)
        {
            using (var workbook = new XLWorkbook(xlsxPath))
            {
                var worksheet = workbook.Worksheet(1);

                // Usando uma matriz para armazenar as células
                var range = worksheet.RangeUsed(); // Faz a leitura de todas as células usadas
                var values = range.Cells().Select(cell => FormatCellValue(cell)).ToArray(); // Convertendo as células para valores formatados

                // Gravar o CSV de uma vez só
                using (var csvStream = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    for (int i = 0; i < range.RowCount(); i++)
                    {
                        var csvLine = string.Join(";", values.Skip(i * range.ColumnCount()).Take(range.ColumnCount()));
                        csvStream.WriteLine(csvLine);
                    }
                }
            }
        }

        private static string FormatCellValue(IXLCell cell)
        {
            if (cell.IsEmpty()) return "";

            string value = cell.Value.ToString();

            // Formata números corretamente
            if (cell.Value.IsNumber)
            {
                value = cell.Value.GetNumber().ToString(CultureInfo.InvariantCulture);
            }

            // Escapa caracteres problemáticos
            if (value.Contains(";") || value.Contains("\""))
            {
                value = $"\"{value.Replace("\"", "\"\"")}\"";
            }

            return value;
        }
    }
}
