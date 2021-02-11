using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace test_all_features_2
{
    class ExcelWorkSheet
    {
        private string workSheetName;
        private DataTable table;

        public ExcelWorkSheet(string workSheetName, DataTable table)
        {
            this.workSheetName = workSheetName;
            this.table = table;
        }

        public string getWorkSheetName()
        {
            return workSheetName;
        }

        public DataTable getTable()
        {
            return table;
        }
    }
}
