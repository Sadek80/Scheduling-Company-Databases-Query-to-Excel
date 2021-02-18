using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using System.Data;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Making the Excel Files
    /// </summary>
    class  HandlingExcelFile
    {
        /// <summary>
        /// Directory Path
        /// </summary>
        private static string path;

        // Fields used to Search for files in the Directory with the same name.
        private string searchPath;
        private string searchName;




        /// <summary>
        /// Public Constructor to Set the Excel File Path with the name Supplied and proper Extension
        /// </summary>
        /// <param name="Name">Excel File Name</param>
        /// <param name="Path">Directory Path</param>
        public HandlingExcelFile(string Name, string Path)
        {
            searchPath = Path;
            path = Path.Replace('/', '\\');
            path = path + "/" + Name + ".xlsx";
            searchName = Name;
        }

        public string getPath()
        {
            return path;
        }

        /// <summary>
        /// Save the Excel File to the Directory Path
        /// </summary>
        /// <param name="workSheets">List of Work Sheets to Add it to the Excel File</param>
        public void makeExcelFile(List<ExcelWorkSheet> workSheets)
        {

            try
            {
                // Using ClosedXML Library XLWorkbook Object
                XLWorkbook workbook = new XLWorkbook();

                using (workbook)
                {

                    // Adding Every Work Sheet of the given List to the Excel File
                    foreach (ExcelWorkSheet workSheet in workSheets)
                        workbook.AddWorksheet(workSheet.getTable(), workSheet.getWorkSheetName());

                    // Searching for Files with the Same Name and Delete it to avoid any Conflict
                    string[] files = System.IO.Directory.GetFiles(searchPath, searchName+"*.xlsx");
                    foreach (string f in files)
                    {
                        System.IO.File.Delete(f);
                    }

                    // Save the File with the given path and name 
                    workbook.SaveAs(path);

                    
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine("Excel File has been Created Successfuly!");
                    Console.ResetColor();

                    workbook.Dispose();


                }
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the HandlingExcelFile: "+ e.Message);
                Console.ResetColor();           
            }
        }

      
    }
}
