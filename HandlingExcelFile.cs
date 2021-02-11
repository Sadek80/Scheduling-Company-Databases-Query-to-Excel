using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using System.Data;

namespace test_all_features_2
{

    class HandlingExcelFile
    {
        public static XLWorkbook workbook = new XLWorkbook();

        private static string path;


        public static void setPath(string Name, string Path)
        {
            path = Path.Replace('/','\\');
            path = path + "/" + Name + ".xlsx";
        }

        public static string getPath()
        {
            return path;
        }
        public static void makeExcelFile(List<ExcelWorkSheet> workSheets)
        {
            try
            {


                using (workbook)
                {


                    foreach (ExcelWorkSheet workSheet in workSheets)
                        workbook.AddWorksheet(workSheet.getTable(), workSheet.getWorkSheetName());

                    workbook.SaveAs(path);
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("in the HandlingExcelFile: "+ e.Message);
            }
        }
    }
}
