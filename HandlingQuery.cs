using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Odbc;
using System.Data;

namespace test_all_features_2
{
    class  HandlingQuery
    {
        public static List<ExcelWorkSheet> workSheets = new List<ExcelWorkSheet>();

        public static void excuteQuery(Query query)
        {
            string connectionString = @"Dsn=Library_System;uid=DESKTOP-N1EUBL2\Sadek";
            using (OdbcConnection connection =
               new OdbcConnection(connectionString))
            {
                connection.Open();

                OdbcCommand command = new OdbcCommand(query.getQuery(), connection);
                connection.Close();

                
                    DataTable table = new DataTable();

                    OdbcDataAdapter adapter = new OdbcDataAdapter(command);

                    adapter.Fill(table);

                    workSheets.Add(new ExcelWorkSheet(query.getQuerySheetName(), table));
                  
            }

            
        
    }

        public static void makeExcelFile()
        {
            HandlingExcelFile.makeExcelFile(workSheets);
        }
    }
}
