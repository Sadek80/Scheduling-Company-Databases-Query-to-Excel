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
       
            // Landing Interface.
            HandlingConsoleInterface.ConnectionStringInterface();

            // Handling all the operation
            HandlingConsoleInterface.MainInterface();


            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Thank you");

            Console.ReadKey();

        }
    }
}
