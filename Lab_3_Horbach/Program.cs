using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Lab_3_Horbach_623p
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> promotions = new List<string>
            {
                "Store A, Buy One Get One Free on Consultation, BOGO123, 2024-05-01, 2024-05-31",
                "Store B, 20% Discount on Cosmetic Procedures, COS20OFF, 2024-05-15, 2024-06-15",
                "Store C, Refer a Friend and Get 25% off Your Next Visit, REF25OFF, 2024-06-01, 2024-06-30"
            };
            List<string> removedPromotions = new List<string>
            {
                "Store D, Special Offer for New Customers, NEWCUST10, 2024-05-20, 2024-06-20",
                "Store E, Summer Sale: 30% Off All Products, SUMMER30, 2024-06-10, 2024-07-10"
            };

            Assembly word_assembly = Assembly.LoadFrom(@"C:\Users\Arina Gorbach\Desktop\Lab_3_Horbach\Library_Word\bin\Debug\net8.0\Library_Word.dll");
            Type word_type = word_assembly.GetType("Library_Word.Class_Word");
            object word_instance = Activator.CreateInstance(word_type);
            MethodInfo word_method = word_type.GetMethod("ExportToWord");
            word_method.Invoke(word_instance, new object[] { promotions, removedPromotions });

            Assembly excel_assembly = Assembly.LoadFrom(@"C:\Users\Arina Gorbach\Desktop\Lab_3_Horbach\Library_Excel\bin\Debug\net8.0\Library_Excel.dll");
            Type excel_type = excel_assembly.GetType("Library_Excel.Class_Excel");
            object excel_instance = Activator.CreateInstance(excel_type);
            MethodInfo excel_method = excel_type.GetMethod("ExportToExcel");
            excel_method.Invoke(excel_instance, new object[] { promotions, removedPromotions });

            Console.WriteLine("Promotion reports have been generated.");
        }
    }
}