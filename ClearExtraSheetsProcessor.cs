using System;
using System.IO;
using System.Linq;
using IronXL;

namespace Specification
{
    public static class ClearExtraSheetsProcessor
    {
        private static readonly string[] PreservedSheets = { "Обработка спецификации", "Дерево узлов Cadmatic Hull", "Для сводных данных", "Условия для объединения", "Название чертежей РКД" };

        public static void ClearSheets(IProgress<int>? progress = null)
        {
            try
            {
                string outputDirectory = @"D:\специфика";
                var files = Directory.GetFiles(outputDirectory, "*.xlsx");
                int totalFiles = files.Length;
                int processedFiles = 0;

                foreach (var file in files)
                {
                    WorkBook workbook = WorkBook.Load(file);
                    var sheetsToDelete = workbook.WorkSheets.Where(ws => !PreservedSheets.Contains(ws.Name)).ToList();
                    foreach (var sheet in sheetsToDelete)
                        workbook.RemoveWorkSheet(sheet.Name);
                    workbook.Save();
                    processedFiles++;
                    progress?.Report((int)((processedFiles / (float)totalFiles) * 100));
                }

                CommonUtilities.ShowSuccess("Лишние листы удалены");
            }
            catch (Exception ex)
            {
                CommonUtilities.ShowError("Произошла ошибка: " + ex.Message);
            }
        }
    }
}