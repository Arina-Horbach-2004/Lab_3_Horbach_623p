using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace Library_Word
{
    public class Class_Word
    {
        public void ExportToWord(List<string> promotions, List<string> removedPromotions)
        {
            // Створення нового екземпляру Word та документа
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();
            Desing_Word(doc, promotions, removedPromotions); // Додавання даних та форматування

            // Шлях до файлу Word
            string wordPath = @"C:\Users\Arina Gorbach\Desktop\PromotionReport.doc";
            doc.SaveAs2(wordPath); // Збереження документа
            doc.Close(); // Закриття документа

            // Звільнення пам'яті та ресурсів COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            wdApp.Quit(); // Закриття Word
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        static void Desing_Word(Word.Document doc, List<string> promotions, List<string> removedPromotions)
        {
            // Додавання заголовка "Promotion Report"
            Word.Paragraph header = doc.Paragraphs.Add();
            header.Range.Text = "Promotion Report";
            header.Range.Font.Size = 24;
            header.Range.Font.Name = "Times New Roman";
            header.Range.InsertParagraphAfter();

            // Додавання заголовка таблиці "Promotions"
            Word.Paragraph promotionsHeader = doc.Paragraphs.Add();
            promotionsHeader.Range.Text = "Active Promotions";
            promotionsHeader.Range.Font.Name = "Times New Roman";
            promotionsHeader.Range.Font.Size = 14;
            promotionsHeader.Range.InsertParagraphAfter();

            // Додавання таблиці з активними промо-акціями
            Word.Table promotionsTable = doc.Tables.Add(promotionsHeader.Range, promotions.Count + 1, 5);
            promotionsTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            promotionsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            promotionsTable.Cell(1, 1).Range.Text = "Store";
            promotionsTable.Cell(1, 2).Range.Text = "Description";
            promotionsTable.Cell(1, 3).Range.Text = "Code";
            promotionsTable.Cell(1, 4).Range.Text = "Start Date";
            promotionsTable.Cell(1, 5).Range.Text = "End Date";
            for (int i = 0; i < promotions.Count; i++)
            {
                string[] promotionData = promotions[i].Split(',');
                for (int j = 0; j < promotionData.Length; j++)
                {
                    promotionsTable.Cell(i + 2, j + 1).Range.Text = promotionData[j].Trim();
                }
            }

            // Додавання заголовка таблиці "Removed Promotions"
            Word.Paragraph removedHeader = doc.Paragraphs.Add();
            removedHeader.Range.Text = "Removed Promotions";
            removedHeader.Range.Font.Name = "Times New Roman";
            removedHeader.Range.Font.Size = 14;
            removedHeader.Range.InsertParagraphAfter();

            // Додавання таблиці з видаленими промо-акціями
            Word.Table removedTable = doc.Tables.Add(removedHeader.Range, removedPromotions.Count + 1, 5);
            removedTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            removedTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            removedTable.Cell(1, 1).Range.Text = "Store";
            removedTable.Cell(1, 2).Range.Text = "Description";
            removedTable.Cell(1, 3).Range.Text = "Code";
            removedTable.Cell(1, 4).Range.Text = "Start Date";
            removedTable.Cell(1, 5).Range.Text = "End Date";
            for (int i = 0; i < removedPromotions.Count; i++)
            {
                string[] promotionData = removedPromotions[i].Split(',');
                for (int j = 0; j < promotionData.Length; j++)
                {
                    removedTable.Cell(i + 2, j + 1).Range.Text = promotionData[j].Trim();
                }
            }
        }
    }
}
