using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ReadExcelAndWriteExcel
{
    class Program
    {
        private List<Town> towns = new List<Town>();
        static void Main(string[] args)
        {
            Console.Write("Give an excel file path: ");
            string path = Console.ReadLine();

            if(File.Exists(path))
            {
                Console.WriteLine("File Exists.");

                Program prg = new Program();

                try
                {
                    File.ReadAllLines(path, Encoding.Default).ToList().ForEach(delegate (string line)
                    {
                        prg.towns.Add(new Town(line));
                    });
                    Console.WriteLine("File Read!");
                    writeExcelFile(prg.towns, "hely2");
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.Message);
                }

            } else
            {
                Console.WriteLine("File does not Exist!");
            }

            Console.ReadKey();
        }
        public static void writeExcelFile<T>(List<T> collection, string filename) where T : Town
        {
            string outPutDir = $"{Directory.GetCurrentDirectory()}/excel";
            if(!Directory.Exists(outPutDir))
            {
                Directory.CreateDirectory(outPutDir);
            }
            try
            {
                Console.WriteLine("Writing File...");
                Application oXL;
                _Workbook oWB;
                _Worksheet oSheet;
                Range oRng;
                object misvalue = System.Reflection.Missing.Value;

                // Start Excel and get Application object.
                oXL = new Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (_Workbook)(oXL.Workbooks.Add(""));
                oSheet = (_Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "name";
                oSheet.Cells[1, 2] = "type";
                oSheet.Cells[1, 3] = "county";
                oSheet.Cells[1, 4] = "district";
                oSheet.Cells[1, 5] = "area";
                oSheet.Cells[1, 6] = "population";
                oSheet.Cells[1, 7] = "apartmentscount";

                //Format A1:G1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "G1").Font.Bold = true;
                oSheet.get_Range("A1", "G1").VerticalAlignment =
                    XlVAlign.xlVAlignCenter;

                for (int i = 0; i < collection.Count; ++i)
                {
                    oSheet.Cells[i + 2, 1] = collection[i].Name;
                    oSheet.Cells[i + 2, 2] = collection[i].Type;
                    oSheet.Cells[i + 2, 3] = collection[i].County;
                    oSheet.Cells[i + 2, 4] = collection[i].District;
                    oSheet.Cells[i + 2, 5] = collection[i].Area;
                    oSheet.Cells[i + 2, 6] = collection[i].Population;
                    oSheet.Cells[i + 2, 7] = collection[i].ApartmentsCount;
                }

                //AutoFit columns A:G.
                oRng = oSheet.get_Range("A1", "G1");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs($"{outPutDir}/{filename}.xls", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
                oXL.Quit(); 
                Console.WriteLine("File Created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
