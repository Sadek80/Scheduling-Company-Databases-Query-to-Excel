using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace test_all_features_2
{
    class HandlingConsoleInterface
    {
        static int operation;

        /// <summary>
        /// this Method enables the user to input the connection string
        /// </summary>
        public static void ConnectionStringInterface()
        {
            Console.WriteLine("Enter your connection string: ");
            string connectionString = Console.ReadLine();

            HandlingQuery.setConnectoinString(connectionString);
        }

        /// <summary>
        /// this Method enables the user to go to query interface or Company Settings
        /// </summary>
        public static void MainInterface()
        {

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("******************************************");
            Console.WriteLine("             Main Interface");
            Console.WriteLine("******************************************");
            Console.WriteLine();

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Create and Send Excel File."
                    + "\n 2. Companies Settings."
                    + "\n 3. Test a Query"
                    + "\n 4. Exit");

                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    CreateAndSendExcelFile();
                }

                else if (operation == 2)
                {
                    CompanySettingsInterface();
                }

                else if (operation == 3)
                {
                    string query;
                    Console.WriteLine();
                    Console.WriteLine("Add your Query: ");
                    query = Console.ReadLine();

                    if (query.Length != 0)
                    {
                        Console.WriteLine("Loading...");
                        HandlingQuery.testQueries(query);
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("you must add a query first.");
                    }
                }

                else if(operation == 4)
                {
                    break;
                }

                else
                    Console.WriteLine("You've Entered a wrong input");
            }
        }



        /// <summary>
        /// this Method enables the user to choose the receiver company with its departments
        /// </summary>
        public static void CreateAndSendExcelFile()
        {
            Console.WriteLine();
            Console.WriteLine();
            string companyName;
            int choice;
            List<string> allCompanies = HandlingLocalDb.ListOfAllCompaniesNames();
            if (allCompanies.Count == 0)
            {
                Console.WriteLine("Please Go to Companies Settings and add companies and departments to the Database");
                Console.WriteLine();
                return;
            }
                Console.WriteLine("Choose the Company you want to send to");
            //TODO list of all companies in the data base to choose from
            
            for (int i = 0; i < allCompanies.Count; i++)
            {
                Console.WriteLine((i+1) + ". " + allCompanies[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

            companyName = allCompanies[choice - 1];
            allCompanies.Clear();
            List<string> departmentToSend = new List<string>();

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine();

                Console.WriteLine("1. Choose the Departments"
                    + "\n 2. Continue to Add a query and Send an Email"
                    + "\n 3. Back to Main Interface");

                operation = int.Parse(Console.ReadLine());
                List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentsNamesInACompany(companyName);

                if (operation == 1)
                {
                    //TODO list of all company departments to choose from
                    if (departmentToSend.Count == 0)
                    {
                        Console.WriteLine();
                        for (int i = 0; i < allDepartments.Count; i++)
                        {
                            Console.WriteLine((i + 1) + ". " + allDepartments[i]);

                        }
                    }
                    else if(departmentToSend.Count == allDepartments.Count)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("You've choosed all the departments");
                        Console.WriteLine();
                        continue;
                    }
                    else
                    {
                        for (int i = 0; i < allDepartments.Count; i++)
                        {
                            Console.WriteLine();
                            foreach(string department in departmentToSend)
                            {
                                if (department != allDepartments[i])
                                {
                                    Console.WriteLine((i + 1) + ". " + allDepartments[i]);

                                }
                            }

                        }
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    departmentToSend.Add(allDepartments[choice - 1]);
                    Console.WriteLine();

                }

                else if(operation == 2)
                {
                    //TODO validation for an empty department list
                    if(departmentToSend.Count == 0)
                    {
                        Console.WriteLine("You Should choose at least one department to continue");
                    }

                    // Go to QueryAndSendEmailInterface
                    else
                    {
                        QueryAndSendEmailInterface(departmentToSend);
                        departmentToSend.Clear();
                        break;
                    }
                   
                }

                else if (operation == 3)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a wrong input.");
            }
        }

        

        /// <summary>
        /// this Method enables the user to handling the company settings
        /// </summary>
        public static void CompanySettingsInterface()
        {
            Console.WriteLine();
            Console.WriteLine();
            while (true)
            {
                Console.WriteLine("1. Add a Company"
                    + "\n 2. Edit a Company"
                    + "\n 3. Delete a Company"
                    + "\n 4. Back to Main Interface");
                operation = int.Parse(Console.ReadLine());

                if(operation == 1)
                {
                    AddCompanyInterface();
                }

                else if(operation == 2)
                {
                    EditCompanyInterface();
                }

                else if(operation == 3)
                {
                    DeleteCompanyInterface();   
                }
                else if (operation == 4)
                {
                    break;
                }

                else
                    Console.WriteLine("You've Entered a wrong input");
                }
        }

        /// <summary>
        /// this method enables the user to add a company
        /// </summary>
        public static void AddCompanyInterface()
        {

            string CompanyName;
            List<Department> departments = new List<Department>();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Enter the Company Name");
            CompanyName = Console.ReadLine();

            while (true)
            {
                string DepartmentName;
                string Email;

                Console.WriteLine();
                Console.WriteLine("1. Add Departments."
                    + "\n 2. Save Company");

                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("Enter Department Name: ");
                    DepartmentName = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Enter Department Email: ");
                    Email = Console.ReadLine();

                    //TODO add the department here with company name on the list of Deaprtment
                    departments.Add(new Department(DepartmentName,Email));

                }

                else if (operation == 2)
                {
                    //TODO if the department list is empty
                    if (departments.Count == 0)
                    {
                        Console.WriteLine("You Should Add Departments First ");
                        continue;
                    }

                    else
                    {
                        HandlingLocalDb.AddCompany(new Company(CompanyName, departments));
                        departments.Clear();
                        break;
                    }
                }

                else
                    Console.WriteLine("You've Entered a wrong input");


            }
        }

        /// <summary>
        /// this method enables the user to edit a company.
        /// </summary>
        public static void EditCompanyInterface()
        {
            string companyName;
            int choice;
            Console.WriteLine();
            Console.WriteLine("Choose a Company to Edit");
            //TODO list of all companies in the data base to choose from
            List<string> allCompanies = HandlingLocalDb.ListOfAllCompaniesNames();
           for(int i = 0; i < allCompanies.Count; i++)
            {
                Console.WriteLine((i+1)+". "+allCompanies[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

            companyName = allCompanies[choice - 1];
            allCompanies.Clear();

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Edit Company Name"
                    + "\n 2. Edit Departments"
                    + "\n 3. Add a new Department"
                    + "\n 4. Back to Companies Settings");
                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    string newName;
                    Console.WriteLine();
                    Console.WriteLine("this the Current Company Name: "+companyName);
                    Console.WriteLine();
                    Console.WriteLine("Enter the new Company Name: ");
                    newName = Console.ReadLine();

                    HandlingLocalDb.UpdateCompanyName(companyName, newName);
                    companyName = newName;
                }

                else if (operation == 2)
                {
                    string departmentName;
                   
                    Console.WriteLine("Choose a Department to Edit");
                    //TODO list of all companies in the data base to choose from
                    List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentsNamesInACompany(companyName);
                    for (int i = 0; i < allDepartments.Count; i++)
                    {
                        Console.WriteLine((i+1) + ". " + allDepartments[i]);
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    departmentName = allDepartments[choice - 1];
                    allDepartments.Clear(); 
                    Console.WriteLine();

                    while (true)
                    {
                        Console.WriteLine();
                        Console.WriteLine("1. Edit Department Name"
                            + "\n 2. Edit Department Email"
                            + "\n 4. Back to Companies Settings");
                        operation = int.Parse(Console.ReadLine());

                        if (operation == 1)
                        {
                            string newDepartmentName;
                            Console.WriteLine();
                            Console.WriteLine("this the Current Department Name: "+ departmentName);
                            Console.WriteLine();
                            Console.WriteLine("Enter the new Department Name: ");
                            newDepartmentName = Console.ReadLine();

                            HandlingLocalDb.UpdateDepartmentName(departmentName,newDepartmentName);
                            break;
                        }
                        else if (operation == 2)
                        {
                            string newDepartmentEmail;
                            Console.WriteLine();
                            Console.WriteLine("this the Current Department Email: "+HandlingLocalDb.getDepartmentEmail(departmentName));
                            Console.WriteLine();
                            Console.WriteLine("Enter the new Department Email: ");
                            newDepartmentEmail = Console.ReadLine();

                            HandlingLocalDb.UpdateDepartmentEmail(departmentName, newDepartmentEmail);
                            break;
                        }
                        else if (operation == 3)
                        {
                            break;
                        }
                        else
                            Console.WriteLine("You've Entered a wrong Input");

                    }
                }
                else if (operation == 3)
                {
                    string DepartmentName;
                    string Email;
                    Console.WriteLine();
                    Console.WriteLine("Enter Department Name: ");
                    DepartmentName = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Enter Department Email: ");
                    Email = Console.ReadLine();

                    //TODO add the department here with company name on the list of Deaprtment
                    HandlingLocalDb.AddDepartment(companyName, new Department(DepartmentName, Email));

                }
                else if (operation == 4)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a wrong Input.");
            }
        }


        /// <summary>
        /// This method enables the user to delete a company
        /// </summary>
        public static void DeleteCompanyInterface()
        {
            int choice;
            string companyName;
            Console.WriteLine();
            Console.WriteLine("Choose a Company to Delete");
            //TODO list of all companies in the data base to choose from
            List<string> allCompanies = HandlingLocalDb.ListOfAllCompaniesNames();
            for (int i = 0; i < allCompanies.Count; i++)
            {
                Console.WriteLine((i+1) + ". " + allCompanies[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

           
                companyName = allCompanies[choice - 1];
                allCompanies.Clear();

          
            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Delete the Company"
                    + "\n 2. Delete a Department"
                    + "\n 3. Back to Companies Settings");
                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " +companyName+ "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        //TODO Delete company
                        HandlingLocalDb.DeleteCompany(companyName);
                        break;
                    }
                    else if (agreement == "no")
                    {
                        break;
                    }
                    else
                        Console.WriteLine("You've Entered a wrong input");
                }

                else if (operation == 2)
                {
                    string departmentName;
                    Console.WriteLine();
                    Console.WriteLine("Choose a Department to delete");
                    //TODO list of all departments in the data base to choose from
                    List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentsNamesInACompany(companyName);
                    for (int i = 0; i < allDepartments.Count; i++)
                    {
                        Console.WriteLine((i+1) + ". " + allDepartments[i]);
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    departmentName = allDepartments[choice - 1];
                    allDepartments.Clear();
                    Console.WriteLine();

                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " +departmentName+ "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        //TODO Delete Department
                        HandlingLocalDb.DeleteDepartment(departmentName);
                        break;
                    }
                    else if (agreement == "no")
                    {
                        break;
                    }
                    else
                        Console.WriteLine("You've Entered a wrong input");
                }

                else if (operation == 3)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a wrong Input");
            }
        }

        /// <summary>
        /// this Method enables the user to add-test query and send an email
        /// </summary>
        public static void QueryAndSendEmailInterface(List<string> Departments)
        {
            List<Query> sql = new List<Query>();

            string query;
            string sheetName;

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.Write("Enter your opertion number: ");
                Console.WriteLine("\n 1. Add sql query " +
                    "\n 2. make the excel file"
                    + "\n 3. Back to Main Interface");
                operation = int.Parse(Console.ReadLine());
                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("Add your Query: ");
                    query = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Add the sheet name for this query: ");
                    sheetName = Console.ReadLine();
                    sql.Add(new Query(query, sheetName));
                }

                else if (operation == 2)
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
                        while (true)
                        {

                        Console.WriteLine("\n 1. Send Email Now"
                            + "\n 2. Send Email in a Specific Time"
                            + "\n 3. Exit");
                        operation = int.Parse(Console.ReadLine());
                        if (operation == 1)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Write email \"Subject\": ");
                            string subject = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write emial \"Body\" text: ");
                            string body = Console.ReadLine();
                            Console.WriteLine();
                            SendingEmail SendingInstance = new SendingEmail(Departments, subject, body, HandlingExcelFile.getPath());
                            Console.WriteLine($"Email with the excel file is Sending....");
                            SendingInstance.sendEmailNow();

                            HandlingQuery.clearList();
                            sql.Clear();
                                break;
                        }
                        else if (operation == 2)
                        {
                            Console.WriteLine();
                            Console.WriteLine("Write email \"Subject\": ");
                            string subject = Console.ReadLine();
                            Console.WriteLine();
                            Console.WriteLine("Write emial \"Body\" text: ");
                            string body = Console.ReadLine();
                            Console.WriteLine();
                            SendingEmail SendingInstance = new SendingEmail(Departments, subject, body, HandlingExcelFile.getPath());

                            Console.WriteLine();
                            Console.WriteLine("Enter Sending Time: ex(16:00), ex(12:30)");
                            string Time = Console.ReadLine();
                            int charLocation = Time.IndexOf(':');
                            int Hour = int.Parse( Time.Substring(0, charLocation));
                            int Minute = int.Parse(Time.Substring(charLocation + 1));

                            DateTime now = DateTime.Now;
                            DateTime firstRun = new DateTime(now.Year, now.Month, now.Day, Hour, Minute, 0, 0);
                            if (now > firstRun)
                            {
                                Console.WriteLine("Time is passed, please Enter valid time.");
                                continue;
                            }
                            TimeSpan timeToGo = firstRun - now;
                            Thread thread = new Thread(() =>
                            {
                                SendingInstance.sendingEmailInTime(timeToGo);

                            });
                            thread.Start();


                            Console.WriteLine("Email with the excel file will be Sent at: "+Hour+":"+""+Minute);

                            HandlingQuery.clearList();
                            sql.Clear();
                                break;
                        }
                        else if(operation == 3)
                            break;
                        else
                            Console.WriteLine("You've entered a wrong input");
                        }

                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("you must add a query first.");
                    }
                }
                else if(operation == 3)
                {
                    break;
                }
                else
                    Console.WriteLine("you've entered a wrong input");
            }
        }

        //public void threadOperations()
        //{

        //}

    }
}
