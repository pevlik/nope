using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace Specification
{
    public static class MaterialDataProcessor
    {
        public static void ProcessMaterialData(string sectionNumber, string projectNumber, IProgress<int>? progress = null)
        {
            try
            {
                string inputFilePath = $@"D:\HullProjects\{projectNumber}\{sectionNumber}\pi\rep-pb_materials_all.list";
                string outputFilePath = $@"D:\специфика\Materials_{sectionNumber}.xlsx";

                if (!File.Exists(inputFilePath)) throw new FileNotFoundException("Файл материалов не найден.");
                CommonUtilities.EnsureDirectoryExists(outputFilePath);

                string[] lines = File.ReadAllLines(inputFilePath, Encoding.UTF8);
                WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                WorkSheet worksheet = workbook.CreateWorkSheet($"Materials_{sectionNumber}");

                int rowIndex = 1;
                foreach (var line in lines)
                {
                    string[] columns = line.Split('|').Select(s => s.Trim()).ToArray();
                    for (int col = 0; col < columns.Length && col < 20; col++)
                        worksheet[$"{(char)('A' + col)}{rowIndex}"].Value = columns[col].Replace(" ", "");
                    rowIndex++;
                    progress?.Report((int)((rowIndex / (float)lines.Length) * 50)); // Прогресс до 50%
                }

                worksheet.RemoveColumn(0); // Удаляем A
                worksheet["A:G"].Replace(" ", "");

                // TODO: Добавить логику из VBA MaterialHandler (SortMaterials, ExcludePipes, ExcludeSqPipes, DelEmpMat, OrderMaterials)
                // Это потребует адаптации циклов VBA для работы с IronXL

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