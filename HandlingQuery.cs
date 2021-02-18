using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Odbc;
using System.Data;
using System.Configuration;
namespace test_all_features_2
{
    /// <summary>
    /// Class for Handling the Execution of the Queries
    /// </summary>
    class HandlingQuery
    {
        /// <summary>
        /// List of the Work Book Work Sheets
        /// </summary>
        private List<ExcelWorkSheet> workSheets = new List<ExcelWorkSheet>();


        public HandlingQuery()
        {

        }

        /// <summary>
        /// Connection String Supplied by the user
        /// </summary>
        static string connectionString;

        /// <summary>
        /// Set the Data Base Connection String
        /// </summary>
        /// <param name="con">Connection String</param>
        public static void setConnectoinString(string con)
        {
            try
            {
                connectionString = con;

            }
            catch (Exception e)
            {
                Console.WriteLine("in Connection String: " + e.Message);
            }
        }

        /// <summary>
        /// Execute a Query and Generate the Data Table that will be Shown in the Query Work Sheet 
        /// </summary>
        /// <param name="query">Query Statement</param>
        public void excuteQuery(Query query)
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

                    // Fill the Table with the Query Result
                    adapter.Fill(table);

                    // Add the Work Sheet to the Work Book  Work Sheets List
                    workSheets.Add(new ExcelWorkSheet(query.getQuerySheetName(), table));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("in the handling query: " + e.Message);
            }
        }

        /// <summary>
        /// Get The Work Sheets
        /// </summary>
        /// <returns>List of Work Sheets</returns>
        public List<ExcelWorkSheet> getWorkSheetsList()
        {
            return workSheets;
        }

        /// <summary>
        /// Clearing the Work Sheet List
        /// </summary>
        public void clearList()
        {
            workSheets.Clear();
        }

        /// <summary>
        /// Method to Test the Data Base Connection String and Showing the Query result in the Console.
        /// </summary>
        /// <param name="query">Query Statemnt</param>
        public void testQueries(string query)
        {
            try
            {

                using (OdbcConnection connection =
                  new OdbcConnection(connectionString))
                {
                    connection.Open();

                    OdbcCommand command = new OdbcCommand(query, connection);
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
                        Console.WriteLine();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("in the test querry: " + e.Message);
            }
        }
    }
}
