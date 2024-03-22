using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace CreateExcelFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Enter the name of the file: ");
            string fileName = Console.ReadLine();

            Console.WriteLine("Enter the column names (separated with ,)");
            string[] columnNames = Console.ReadLine().Split(',');

            string repoParh = @"C:\Path\To\Your\Repo";
            string filePath = Path.Combine(repoParh, fileName + ".xlsx");


            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < columnNames.Length; i++)
                {
                    worksheet.Cells[1, i +1].Value = columnNames[i].Trim();
                }

                int row = 2;
                while (true)
                {
                    Console.WriteLine("Enter the values for the row " + row + " separated with ' , ' or press 'q' to exit!");
                    string input = Console.ReadLine();
                    if (input.ToLower() == "q")
                        break;


                    string[] values = input.Split(',');
                    for(int i = 0; i < values.Length && i < columnNames.Length; i++)
                    {
                        worksheet.Cells[row, i + 1].Value = values[i].Trim();
                    }
                    row++;
                }

                FileInfo file = new FileInfo(filePath);
                excelPackage.SaveAs(file);
            }

            Console.WriteLine("Excel file was created!");
            Console.ReadLine(); 
        }
    }
}
