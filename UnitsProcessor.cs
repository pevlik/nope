using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace Specification
{
    public static class UnitsProcessor
    {
        public static void ProcessUnits(string sectionNumber, string projectNumber, IProgress<int>? progress = null)
        {
            try
            {
                string inputFilePath = $@"D:\HullProjects\{projectNumber}\{sectionNumber}\pi\rep-pb_units_all.list";
                string outputFilePath = $@"D:\специфика\Units_{sectionNumber}.xlsx";

                if (!File.Exists(inputFilePath)) throw new FileNotFoundException("Файл узлов не найден.");
                CommonUtilities.EnsureDirectoryExists(outputFilePath);

                string[] lines = File.ReadAllLines(inputFilePath, Encoding.UTF8);
                WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                WorkSheet worksheet = workbook.CreateWorkSheet($"Units_{sectionNumber}");

                int rowIndex = 1;
                foreach (var line in lines)
                {
                    string[] columns = line.Split('|').Select(s => s.Trim()).ToArray();
                    for (int col = 0; col < columns.Length && col < 20; col++)
                        worksheet[$"{(char)('A' + col)}{rowIndex}"].Value = columns[col].Replace(" ", "");
                    rowIndex++;
                    progress?.Report((int)((rowIndex / (float)lines.Length) * 100));
                }

                worksheet.RemoveColumn(0); // Удаляем A
                worksheet["A:G"].Replace(" ", "");

                // TODO: Добавить дополнительную обработку узлов, если требуется (аналог VBA UnitsHandler)

                workbook.SaveAs(outputFilePath);
                CommonUtilities.ShowSuccess($"Файл успешно сохранен: {outputFilePath}");
            }
            catch (Exception ex)
            {
                CommonUtilities.ShowError("Произошла ошибка: " + ex.Message);
            }
        }
    }
}