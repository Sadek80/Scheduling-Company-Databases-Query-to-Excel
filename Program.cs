using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Configuration;

namespace test_all_features_2
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Query> sql = new List<Query>();

            int operation;
            string query;
            string sheetName;

            Console.WriteLine("Enter your connection string: ");
            string connectionString = Console.ReadLine();

            HandlingQuery.setConnectoinString(connectionString);

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.Write("Enter your opertion number: ");
                Console.WriteLine("\n 1. Add sql query " +
                    "\n 2. make the excel file"
                    +"\n 3. Test Query");
                operation = int.Parse(Console.ReadLine());
                if(operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("Add your Query: ");
                    query = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Add the sheet name for this query: ");
                    sheetName = Console.ReadLine();
                    sql.Add(new Query(query,sheetName));
                }

                else if(operation == 2)
                {
                    if (sql.Count != 0)
                    {


                        Console.WriteLine();
                        Console.WriteLine("Enter the Excel file name (without extension): ");
                        string name = Console.ReadLine();
                        Console.WriteLine();
                        Console.WriteLine("Enter the file path: ");
                        string path = Console.ReadLine();
                        HandlingExcelFile.setPath(name, path);


                        foreach (Query sqlQuery in sql)
                        {
                            HandlingQuery.excuteQuery(sqlQuery);
                        }

                        HandlingQuery.makeExcelFile();
                        Console.WriteLine("\n 1. Send Email"
                            + "\n 2. Exit");
                        operation = int.Parse(Console.ReadLine());
                        if (operation == 1)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Write \"From\" emial: ");
                            string from = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write \"From\" emial password: ");
                            string password = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write \"To\" emial: ");
                            string to = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write email \"Subject\": ");
                            string subject = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write emial \"Body\" text: ");
                            string body = Console.ReadLine();
                            Console.WriteLine();
                            SendingEmail.setEmailData(from, password, to, subject, body, HandlingExcelFile.getPath());
                            Console.WriteLine($"Email with the excel file is Sending....");
                            SendingEmail.sendEmail();

                            HandlingQuery.clearList();
                            sql.Clear();
                        }
                        else
                            break;
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("you must add a query first.");
                    }
                }
                else if(operation == 3)
                {
                    if (sql.Count != 0)
                    {


                        foreach (Query sqlQuery in sql)
                        {
                            HandlingQuery.testQueries(sqlQuery);

                        }
                        HandlingQuery.clearList();
                        sql.Clear();
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("you must add a query first.");
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Thank you");

        }
    }
}
