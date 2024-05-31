using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace Library_Excel
{
    public class Class_Excel
    {
        public void ExportToExcel(List<string> promotions, List<string> removedPromotions)
        {
            // Створення нового екземпляра Excel та робочої книги
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlBook = xlApp.Workbooks.Add();
            Excel.Worksheet xlSheet1 = xlBook.Sheets[1];
            xlSheet1.Name = "Promotion"; // Назва першого аркуша
            Desing_Excel(xlSheet1, promotions, removedPromotions); // Призначення даних та форматування

            // Шлях до файлу Excel
            string excelPath = @"C:\Users\Arina Gorbach\Desktop\PromotionReport.xlsx";
            xlBook.SaveAs(excelPath); // Збереження книги
            xlBook.Close(); // Закриття книги
            xlApp.Quit(); // Закриття Excel після збереження книги

            // Звільнення пам'яті та ресурсів COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        static void Desing_Excel(Excel.Worksheet xlSheet1, List<string> promotions, List<string> removedPromotions)
        {
            // Заповнення першого аркуша з активними промо-акціями
            xlSheet1.Cells[1, 1].Value = "Магазин";
            xlSheet1.Cells[1, 2].Value = "Опис";
            xlSheet1.Cells[1, 3].Value = "Код";
            xlSheet1.Cells[1, 4].Value = "Дата початку";
            xlSheet1.Cells[1, 5].Value = "Дата закінчення";
            for (int i = 0; i < promotions.Count; i++)
            {
                string[] promotionData = promotions[i].Split(',');
                for (int j = 0; j < promotionData.Length; j++)
                {
                    xlSheet1.Cells[i + 2, j + 1].Value = promotionData[j].Trim();
                }
            }

            // Заповнення другого аркуша з видаленими промо-акціями
            Excel.Worksheet xlSheet2 = xlSheet1.Parent.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing) as Excel.Worksheet;
            xlSheet2.Name = "Видалені промо-акції";
            xlSheet2.Cells[1, 1].Value = "Магазин";
            xlSheet2.Cells[1, 2].Value = "Опис";
            xlSheet2.Cells[1, 3].Value = "Код";
            xlSheet2.Cells[1, 4].Value = "Дата початку";
            xlSheet2.Cells[1, 5].Value = "Дата закінчення";
            for (int i = 0; i < removedPromotions.Count; i++)
            {
                string[] promotionData = removedPromotions[i].Split(',');
                for (int j = 0; j < promotionData.Length; j++)
                {
                    xlSheet2.Cells[i + 2, j + 1].Value = promotionData[j].Trim();
                }
            }

            // Застосування форматування та автоматичне налаштування ширини стовпців для першого аркуша
            Excel.Range tableRange1 = xlSheet1.Range["A1", xlSheet1.Cells[promotions.Count + 1, 5]];
            tableRange1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange1.Font.Name = "Times New Roman";
            tableRange1.Font.Size = 14;
            tableRange1.Columns.AutoFit();

            // Застосування форматування та автоматичне налаштування ширини стовпців для другого аркуша
            Excel.Range tableRange2 = xlSheet2.Range["A1", xlSheet2.Cells[removedPromotions.Count + 1, 5]];
            tableRange2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange2.Font.Name = "Times New Roman";
            tableRange2.Font.Size = 14;
            tableRange2.Columns.AutoFit();
        }
    }
}
