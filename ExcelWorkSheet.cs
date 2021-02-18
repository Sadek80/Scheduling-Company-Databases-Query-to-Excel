using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Organizing the Excel File Work Sheet
    /// </summary>
    class ExcelWorkSheet
    {
        /// <summary>
        /// Work Sheet Name
        /// </summary>
        private string workSheetName;

        /// <summary>
        /// Data Table that represented in the Work Sheet
        /// </summary>
        private DataTable table;

        /// <summary>
        /// Public Constructor For Setting Up the Work Seer Default Data
        /// </summary>
        /// <param name="workSheetName"></param>
        /// <param name="table"></param>
        public ExcelWorkSheet(string workSheetName, DataTable table)
        {
            this.workSheetName = workSheetName;
            this.table = table;
        }


        /// <summary>
        /// Get the Work Sheet Name
        /// </summary>
        /// <returns>Work Sheet Name</returns>
        public string getWorkSheetName()
        {
            return workSheetName;
        }

        /// <summary>
        /// Get Data Table Shown in the Work Sheet
        /// </summary>
        /// <returns>Data Table</returns>
        public DataTable getTable()
        {
            return table;
        }
    }
}
