using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;
using IronXL.Styles;

namespace Specification
{
    public static class Specification
    {
        public static void ProcessSpecification(string sectionNumber, string projectNumber, IProgress<int>? progress = null)
        {
            try
            {
                string inputFilePath = $@"D:\HullProjects\{projectNumber}\{sectionNumber}\pi\rep-specification.list";
                string outputDirectory = @"D:\специфика";
                string outputFilePath = Path.Combine(outputDirectory, $"Spec_{sectionNumber}.xlsx");

                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException("Файл спецификации не найден.");

                string[] lines = File.ReadAllLines(inputFilePath, Encoding.UTF8);
                WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                WorkSheet worksheet = workbook.CreateWorkSheet($"Spec_{sectionNumber}");

                ProcessDataToExcel(worksheet, lines, sectionNumber, progress);

                workbook.SaveAs(outputFilePath);
                MessageBox.Show($"Файл успешно сохранен: {outputFilePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ProcessDataToExcel(WorkSheet worksheet, string[] lines, string sectionNumber, IProgress<int>? progress)
        {
            int rowIndex = 2; // Начинаем со второй строки
            int amountSubsections = 80;
            string[] cyrillicLetters = { "А", "Б", "В", "Г", "Д", "Е", "Ж", "И", "К", "Л", "М", "Н", "П", "Р", "С", "Т", "У", "Ф", "Ц", "Ш", "Щ", "Э", "Ю", "Я" };

            // Заполняем данные
            foreach (var line in lines)
            {
                string[] columns = line.Split('|').Select(s => s.Trim()).ToArray();
                for (int col = 0; col < columns.Length && col < 20; col++) // До Q
                {
                    worksheet[$"{(char)('A' + col)}{rowIndex}"].Value = columns[col].Replace(" ", "");
                }
                rowIndex++;
            }

            // Удаляем столбец A и очищаем M
            worksheet.RemoveColumn(0); // Удаляем A
            worksheet["L:L"].ClearContents(); // M стал L

            // Применяем начальное форматирование
            var rangeAP = worksheet["A:P"];
            rangeAP.Style.Font.Name = "Arial";
            rangeAP.Style.Font.Height = 10;
            rangeAP.Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
            worksheet["A:A"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center; // B стал A
            worksheet["F:F"].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center; // G стал F

            // Границы для B:M (A:L после удаления A)
            var rangeBM = worksheet["A:L"];
            rangeBM.Style.LeftBorder.Type = BorderType.Medium;
            rangeBM.Style.RightBorder.Type = BorderType.Medium;
            rangeBM.Style.TopBorder.Type = BorderType.Thin;
            rangeBM.Style.BottomBorder.Type = BorderType.Thin;
            rangeBM.Style.LeftBorder.Type = BorderType.DashDotDot;

            // Границы для N:Q (M:P после удаления A)
            var rangeNQ = worksheet["M:P"];
            rangeNQ.Style.TopBorder.Type = BorderType.DashDotDot; // Пунктир как xlDot
            rangeNQ.Style.BottomBorder.Type = BorderType.DashDotDot;
            rangeNQ.Style.RightBorder.Type = BorderType.DashDotDot;
            rangeNQ.Style.LeftBorder.Type = BorderType.DashDotDot;

            // Формат чисел для K:L (J:K после удаления A)
            worksheet["J:J"].FormatString = "0.0";
            worksheet["K:K"].FormatString = "0.0";

            // Обработка подсекций
            rowIndex = 2;
            for (int x = 1; x <= amountSubsections; x++)
            {
                string subsecName = x < 10 ? $"Узел:S0{x}" : $"Узел:S{x}";
                string subsecNumber = GetSubsectionNumber(x, cyrillicLetters);

                var plCell = worksheet["B:B"].FirstOrDefault(cell => cell.Value?.ToString().Contains(subsecName + "PL") == true);
                if (plCell != null)
                {
                    int insertRow = plCell.RowIndex;
                    worksheet.InsertRow(insertRow);
                    worksheet[$"B{insertRow}"].Value = $"Подсекция {subsecNumber}";
                    worksheet[$"B{insertRow}"].Style.Font.Bold = true;
                    rowIndex = insertRow + 1;

                    var unitCell = worksheet["B:B"].FirstOrDefault(cell => cell.Value?.ToString().Contains(subsecName + "_") == true);
                    if (unitCell != null)
                    {
                        insertRow = unitCell.RowIndex;
                        worksheet.InsertRow(insertRow);
                        worksheet[$"B{insertRow}"].Value = $"Узлы на подсекцию {subsecNumber}";
                        rowIndex = insertRow + 1;
                    }

                    worksheet["B:B"].Replace(subsecName + "PL", "Листы настила");
                    worksheet["B:B"].Replace(subsecName + "PR", "Ребра жесткости настила");
                    worksheet["B:B"].Replace(subsecName + "_R", $"Россыпь на подсекцию {subsecNumber}");
                    worksheet["B:B"].Replace(subsecName + "_00", "Узел №");
                    worksheet["B:B"].Replace(subsecName + "_0", "Узел №");
                    worksheet["B:B"].Replace(subsecName + "_", "Узел №");
                }
                else
                {
                    var unitCell = worksheet["B:B"].FirstOrDefault(cell => cell.Value?.ToString().Contains(subsecName + "_") == true);
                    if (unitCell != null)
                    {
                        int insertRow = unitCell.RowIndex;
                        worksheet.InsertRow(insertRow);
                        worksheet[$"B{insertRow}"].Value = $"Подсекция {subsecNumber}";
                        worksheet[$"B{insertRow}"].Style.Font.Bold = true;
                        rowIndex = insertRow + 1;

                        worksheet["B:B"].Replace(subsecName + "_R", $"Россыпь на подсекцию {subsecNumber}");
                        worksheet["B:B"].Replace(subsecName + "_00", "Узел №");
                        worksheet["B:B"].Replace(subsecName + "_0", "Узел №");
                        worksheet["B:B"].Replace(subsecName + "_", "Узел №");
                    }
                }

                progress?.Report((int)((x / (float)amountSubsections) * 100));
            }

            // Узлы на секцию/стапель/плаву
            ProcessUnits(worksheet, "SR", "на секцию");
            ProcessUnits(worksheet, "SS", "на стапель");
            ProcessUnits(worksheet, "ST", "на плаву");

            // Задаем имена деталей в столбце D (C после удаления A)
            NameColumn(worksheet);

            // Устанавливаем ширину столбцов через свойства Column
            worksheet.Columns[0].Width = 60;   // A (был B)
            worksheet.Columns[1].Width = 5;   // B (был C)
            worksheet.Columns[2].Width = 23;  // C (был D)
            worksheet.Columns[3].Width = 24;  // D (был E)
            worksheet.Columns[4].Width = 9;   // E (был F)
            worksheet.Columns[5].Width = 10;  // F (был G)
            worksheet.Columns[6].Width = 6;   // G (был H)
            worksheet.Columns[7].Width = 7;   // H (был I)
            worksheet.Columns[8].Width = 8;   // I (был J)
            worksheet.Columns[9].Width = 8;   // J (был K)
            worksheet.Columns[10].Width = 7;  // K (был L)
            worksheet.Columns[11].Width = 24; // L (был M)
            worksheet.Columns[12].Width = 11; // M (был N)
            worksheet.Columns[13].Width = 11; // N (был O)
            worksheet.Columns[14].Width = 11; // O (был P)
            worksheet.Columns[15].Width = 10; // P (был Q)

            // Удаляем столбцы R:T (Q:S после удаления A, индексы 16-18)
            int columnCount = worksheet.Columns.Count();
            worksheet.RemoveColumn(15); // Последний столбец
            worksheet.RemoveColumn(15); // Предпоследний
            worksheet.RemoveColumn(15); // Предпредпоследний

            // Добавляем итоги
            AddFooter(worksheet);
        }

        private static void ProcessUnits(WorkSheet worksheet, string unitType, string description)
        {
            var unitCell = worksheet["B:B"].FirstOrDefault(cell => cell.Value?.ToString().Contains($"Узел:{unitType}_") == true);
            if (unitCell != null)
            {
                if (unitCell.Value?.ToString() == $"Узел:{unitType}_R")
                {
                    worksheet["B:B"].Replace($"Узел:{unitType}_R", $"Россыпь {description}");
                }
                else
                {
                    int row = unitCell.RowIndex;
                    worksheet.InsertRow(row);
                    worksheet[$"B{row}"].Value = $"Узлы {description}";
                    worksheet[$"B{row}"].Style.Font.Bold = true;

                    worksheet["B:B"].Replace($"Узел:{unitType}_R", $"Россыпь {description}");
                    worksheet["B:B"].Replace($"Узел:{unitType}_00", "Узел №");
                    worksheet["B:B"].Replace($"Узел:{unitType}_0", "Узел №");
                    worksheet["B:B"].Replace($"Узел:{unitType}_", "Узел №");
                }
            }
        }

        private static void NameColumn(WorkSheet worksheet)
        {
            var rows = worksheet["A:P"].Where(cell => cell.RowIndex > 1 && !string.IsNullOrEmpty(cell.Value?.ToString()));
            foreach (var cell in rows.Where(c => c.ColumnIndex == 4)) // E (был F)
            {
                int row = cell.RowIndex;
                string fValue = cell.Value?.ToString() ?? ""; // E
                string iValue = worksheet[$"H{row}"].Value?.ToString() ?? ""; // H (был I)
                string jValue = worksheet[$"I{row}"].Value?.ToString() ?? ""; // I (был J)
                string rValue = worksheet[$"P{row}"].Value?.ToString() ?? ""; // P (был R)

                if (fValue == "PL") worksheet[$"C{row}"].Value = "Лист";
                else if (fValue == "Bkt") worksheet[$"C{row}"].Value = "Кница";
                else if (fValue == "ST") worksheet[$"C{row}"].Value = "Заделка";
                else if (fValue == "FB") worksheet[$"C{row}"].Value = $"Полоса {iValue}X{jValue}";
                else if (fValue == "P")
                {
                    if (jValue == "8" && iValue == "80") worksheet[$"C{row}"].Value = "Полособульб №8";
                    else if (jValue == "6" && iValue == "100") worksheet[$"C{row}"].Value = "Полособульб №10";
                    else if (jValue == "7" && iValue == "120") worksheet[$"C{row}"].Value = "Полособульб №12";
                    else if (jValue == "7" && iValue == "140") worksheet[$"C{row}"].Value = "Полособульб №14а";
                    else if (jValue == "9" && iValue == "140") worksheet[$"C{row}"].Value = "Полособульб №14б";
                    else if (jValue == "8" && iValue == "160") worksheet[$"C{row}"].Value = "Полособульб №16а";
                    else if (jValue == "10" && iValue == "160") worksheet[$"C{row}"].Value = "Полособульб №16б";
                    else if (jValue == "9" && iValue == "180") worksheet[$"C{row}"].Value = "Полособульб №18а";
                    else worksheet[$"C{row}"].Value = $"Полособульб {iValue}X{jValue}";
                }
                else if (fValue == "AS") worksheet[$"C{row}"].Value = $"Круг {jValue}";
                else if (fValue == "PY" && rValue.Length > 5) worksheet[$"C{row}"].Value = $"Труба круглая {rValue.Substring(5)}";
                else if (fValue == "KO" && rValue.Length > 7) worksheet[$"C{row}"].Value = $"Труба квадратная {rValue.Substring(7)}";
            }
        }

        private static void AddFooter(WorkSheet worksheet)
        {
            var firstCell = worksheet["J:J"].First(cell => !string.IsNullOrEmpty(cell.Value?.ToString())); // J (был K)
            var lastCell = worksheet["J:J"].Last(cell => !string.IsNullOrEmpty(cell.Value?.ToString()));

            int footerRow = lastCell.RowIndex + 3;
            worksheet[$"D{footerRow}"].Value = "Масса деталей"; // D (был E)
            worksheet[$"J{footerRow}"].Formula = $"=SUM(J{firstCell.RowIndex}:J{lastCell.RowIndex})";
            worksheet[$"D{footerRow + 1}"].Value = "Масса с наплавленным металлом";
            worksheet[$"J{footerRow + 1}"].Formula = $"=J{footerRow}*1.01";
        }

        private static string GetSubsectionNumber(int x, string[] cyrillicLetters)
        {
            int index = (x - 1) % cyrillicLetters.Length;
            int counter = (x - 1) / cyrillicLetters.Length + 1;
            return x <= cyrillicLetters.Length ? cyrillicLetters[x - 1] : cyrillicLetters[index] + counter;
        }
    }
}