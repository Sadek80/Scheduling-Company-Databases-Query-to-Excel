using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Odbc;
using System.Data;
using System.Configuration;
namespace test_all_features_2
{
    class  HandlingQuery
    {
        public static List<ExcelWorkSheet> workSheets = new List<ExcelWorkSheet>();




        static string connectionString;

        public static void setConnectoinString(string con)
        {
            connectionString = con;
        }

        public static void excuteQuery(Query query)
        {
          try
          {
                
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
            catch (Exception e)
            {
                Console.WriteLine("in the handling query: "+e.Message);
            }





        }

        public static void makeExcelFile()
        {
            HandlingExcelFile.makeExcelFile(workSheets);
        }

        public static void clearList()
        {
            workSheets.Clear();
        }

        public static void testQueries(Query query)
        {
            try
            {

                using (OdbcConnection connection =
                  new OdbcConnection(connectionString))
                {
                    connection.Open();

                    OdbcCommand command = new OdbcCommand(query.getQuery(), connection);
                    using (var reader = command.ExecuteReader())
                    {
                        Console.WriteLine("Query result:");
                        // Print column names
                        var sbCol = new System.Text.StringBuilder();
                        for (var i = 0; i < reader.FieldCount; i++)
                        {
                            sbCol.Append(reader.GetName(i).PadRight(20));
                        }
                        Console.WriteLine(sbCol.ToString());
                        // Print rows
                        while (reader.Read())
                        {
                            var sbRow = new System.Text.StringBuilder();
                            for (var i = 0; i < reader.FieldCount; i++)
                            {
                                sbRow.Append(reader[i].ToString().PadRight(20));
                            }
                            Console.WriteLine(sbRow.ToString());
                        }
                        connection.Close();

                    }
                }
           }
            catch(Exception e)
            {
                Console.WriteLine("in the handling querry: "+e.Message);
            }
        }
    }
}
