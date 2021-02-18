using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace test_all_features_2
{
    /// <summary>
    /// This Class for Handling the Console Interface and Controling the Navigation Through the program.
    /// </summary>
    class HandlingConsoleInterface
    {
        static int operation;


        /// <summary>
        /// this Method enables the user to Enter the connection string
        /// </summary>
        public static void ConnectionStringInterface()
        {
            string connectionString;
            while (true)
            {
                Console.WriteLine("Enter your connection string: ");
                connectionString = Console.ReadLine();
                if (connectionString == "")
                {
                    continue;
                }
                else
                {
                    break;
                }

            }


            HandlingQuery.setConnectoinString(connectionString);
        }

        /// <summary>
        /// this Method enables the user to Navigate through the program
        /// </summary>
        public static void MainInterface()
        {
            try
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
                    Console.WriteLine("1. Create and Send Excel File Now."
                        + "\n 2. Settings."
                        + "\n 3. Test a Query"
                        + "\n 4. Run a Schedule Task (use it if the Program was Turned off)"
                        + "\n 5. Stop a Schedule Task"
                        + "\n 6. Exit");

                    operation = int.Parse(Console.ReadLine());

                    if (operation == 1)
                    {
                        CreateAndSendExcelFile();
                    }

                    else if (operation == 2)
                    {
                        SettingsInterface();
                    }

                    else if (operation == 3)
                    {
                        string query;
                        Console.WriteLine();
                        Console.WriteLine("Add your Query: ");
                        query = Console.ReadLine();

                        HandlingQuery handlingQuery = new HandlingQuery();
                        if (query.Length != 0)
                        {
                            Console.WriteLine("Loading...");

                            handlingQuery.testQueries(query);
                        }
                        else
                        {
                            Console.WriteLine();
                            Console.WriteLine("you must add a query first.");
                        }
                    }

                    else if (operation == 4)
                    {
                        List<string> Tasks = HandlingLocalDb.ListOfAllTasks();
                        int choice;
                        Console.WriteLine();
                        if (Tasks.Count == 0)
                        {
                            Console.WriteLine("There is no Schedule Tasks to run, Please Add one First");
                        }
                        else
                        {
                            for (int i = 0; i < Tasks.Count; i++)
                            {
                                Console.WriteLine((i + 1) + ". " + Tasks[i]);
                            }

                            Console.WriteLine();
                            choice = int.Parse(Console.ReadLine());

                            string taskName = Tasks[choice - 1];
                            Tasks.Clear();
                            HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                            handlingLocalDb.runTask(HandlingLocalDb.getTaskByName(taskName));
                        }

                    }

                    else if (operation == 5)
                    {
                        try
                        {

                            List<string> Tasks = HandlingLocalDb.ListOfAllTasks();
                            int choice;
                            Console.WriteLine();
                            if (Tasks.Count == 0)
                            {
                                Console.WriteLine("There is no Schedule Tasks to run, Please Add one First");
                            }
                            else
                            {
                                for (int i = 0; i < Tasks.Count; i++)
                                {
                                    Console.WriteLine((i + 1) + ". " + Tasks[i]);
                                }

                                Console.WriteLine();
                                choice = int.Parse(Console.ReadLine());

                                string taskName = Tasks[choice - 1];
                                Tasks.Clear();
                                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                                handlingLocalDb.AbortTaskThread(taskName);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("in the kill task: " + e.Message);
                        }
                    }

                    else if (operation == 6)
                    {
                        break;
                    }
                    else
                        Console.WriteLine("You've Entered a wrong input");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("in the main Inerface: " + e.Message);
            }
        }



        /// <summary>
        /// Method enables the user to create and send excel file now
        /// </summary>
        public static void CreateAndSendExcelFile()
        {
            try
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
                    Console.WriteLine((i + 1) + ". " + allCompanies[i]);
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
                        else if (departmentToSend.Count == allDepartments.Count)
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
                                foreach (string department in departmentToSend)
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

                    else if (operation == 2)
                    {
                        //TODO validation for an empty department list
                        if (departmentToSend.Count == 0)
                        {
                            Console.WriteLine("You Should choose at least one department to continue");
                        }

                        // Go to QueryAndSendEmailInterface
                        else
                        {
                            QueryAndSendEmailInterface(departmentToSend, companyName);
                            departmentToSend.Clear();
                            allDepartments.Clear();
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
            catch (Exception e)
            {
                Console.WriteLine("in the Creating File Now: " + e.Message);
            }
        }



        /// <summary>
        /// Method enables the user to Edit the Companies and Schedule tasks settings
        /// </summary>
        public static void SettingsInterface()
        {
            try
            {

                Console.WriteLine();
                Console.WriteLine();
                while (true)
                {
                    Console.WriteLine("1. Company Settings"
                        + "\n 2. Schedule Task Settings"
                        + "\n 3. Back To Main Interface");

                    operation = int.Parse(Console.ReadLine());

                    if (operation == 1)
                    {
                        CompanySettingsInterface();
                    }

                    else if (operation == 2)
                    {
                        //TODO schedule function
                        ScheduleTaskInterface();
                    }

                    else if (operation == 3)
                    {
                        break;
                    }
                    else
                        Console.WriteLine("You've Entered a Wrong Input");


                }
            }
            catch (Exception e)
            {
                Console.WriteLine("in the settings: " + e.Message);
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
                    + "\n 4. Back to Settings");
                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    AddCompanyInterface();
                }

                else if (operation == 2)
                {
                    EditCompanyInterface();
                }

                else if (operation == 3)
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
            Console.WriteLine("Enter the Company Name (Must Be Unique): ");
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
                    departments.Add(new Department(DepartmentName, Email));

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
            for (int i = 0; i < allCompanies.Count; i++)
            {
                Console.WriteLine((i + 1) + ". " + allCompanies[i]);
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
                    Console.WriteLine("this the Current Company Name: " + companyName);
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
                        Console.WriteLine((i + 1) + ". " + allDepartments[i]);
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
                            Console.WriteLine("this the Current Department Name: " + departmentName);
                            Console.WriteLine();
                            Console.WriteLine("Enter the new Department Name: ");
                            newDepartmentName = Console.ReadLine();

                            HandlingLocalDb.UpdateDepartmentName(departmentName, newDepartmentName, companyName);
                            break;
                        }
                        else if (operation == 2)
                        {
                            string newDepartmentEmail;
                            Console.WriteLine();
                            Console.WriteLine("this the Current Department Email: " + HandlingLocalDb.getDepartmentEmail(departmentName, companyName));
                            Console.WriteLine();
                            Console.WriteLine("Enter the new Department Email: ");
                            newDepartmentEmail = Console.ReadLine();

                            HandlingLocalDb.UpdateDepartmentEmail(departmentName, newDepartmentEmail, companyName);
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
                Console.WriteLine((i + 1) + ". " + allCompanies[i]);
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
                    Console.WriteLine(" Are you sure you want to delet the " + companyName + "? (yes) or (no)");
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
                        Console.WriteLine((i + 1) + ". " + allDepartments[i]);
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    departmentName = allDepartments[choice - 1];
                    allDepartments.Clear();
                    Console.WriteLine();

                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " + departmentName + "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        //TODO Delete Department
                        HandlingLocalDb.DeleteDepartment(departmentName, companyName);
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
        /// this Method enables the user to add the query and set the Excel file Information to send now
        /// </summary>
        public static void QueryAndSendEmailInterface(List<string> Departments, string companyName)
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
                        HandlingExcelFile handlingExcelFile = new HandlingExcelFile(name, path);

                        HandlingQuery handlingQuery = new HandlingQuery();
                        foreach (Query sqlQuery in sql)
                        {

                            handlingQuery.excuteQuery(sqlQuery);
                        }


                        handlingExcelFile.makeExcelFile(handlingQuery.getWorkSheetsList());
                        while (true)
                        {

                            Console.WriteLine("\n 1. Send Email Now"
                                + "\n 2. Exit");
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
                                SendingEmail SendingInstance = new SendingEmail(Departments, subject, body, handlingExcelFile.getPath(), companyName);
                                Console.WriteLine($"Email with the excel file is Sending....");
                                SendingInstance.sendEmailNow();

                                handlingQuery.clearList();
                                sql.Clear();
                                break;
                            }
                            else if (operation == 2)
                            {
                                break;
                            }
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
                else if (operation == 3)
                {
                    break;
                }
                else
                    Console.WriteLine("you've entered a wrong input");
            }
        }

        /// <summary>
        /// Method enables the user to Handling Schedule Tasks
        /// </summary>
        public static void ScheduleTaskInterface()
        {
            Console.WriteLine();
            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Add Schedule Task"
                    + "\n 2. Edit Schedule Task"
                    + "\n 3. Delete Schedule Task"
                    + "\n 4. Back to Settings");

                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    //Add task function
                    AddScheduleTask();
                }

                else if (operation == 2)
                {
                    //Edit Task Fuction
                    EditScheduleTask();
                }

                else if (operation == 3)
                {
                    //Delete TASK Gunction
                    DeleteScheduleTask();
                }

                else if (operation == 4)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a Wrong Input");
            }
        }

        /// <summary>
        /// Enables the user to Add task and start running it in the background
        /// </summary>
        public static void AddScheduleTask()
        {
            int choice;
            string companyName;
            Console.WriteLine();
            Console.WriteLine("Enter Schedule Task Name (Must Be Unique): ");
            string scheduleName = Console.ReadLine();

            //List of all Companies
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
                Console.WriteLine((i + 1) + ". " + allCompanies[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

            companyName = allCompanies[choice - 1];
            allCompanies.Clear();
            List<Department> scheduledEmailToSend = new List<Department>();
            while (true)
            {
                Console.WriteLine("1. Choose the Departments"
                      + "\n 2. Continue to Add a Query"
                      + "\n 3. Back to Schedule Settings");

                operation = int.Parse(Console.ReadLine());
                List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentsNamesInACompany(companyName);

                if (operation == 1)
                {
                    //TODO list of all company departments to choose from
                    if (scheduledEmailToSend.Count == 0)
                    {
                        Console.WriteLine();
                        for (int i = 0; i < allDepartments.Count; i++)
                        {
                            Console.WriteLine((i + 1) + ". " + allDepartments[i]);

                        }
                    }
                    else if (scheduledEmailToSend.Count == allDepartments.Count)
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
                            foreach (Department department in scheduledEmailToSend)
                            {
                                if (department.getName() != allDepartments[i])
                                {
                                    Console.WriteLine((i + 1) + ". " + allDepartments[i]);

                                }
                            }

                        }
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    string departmentName = allDepartments[choice - 1];
                    Console.WriteLine();
                    allDepartments.Clear();

                    Console.WriteLine("Enter the Time you want to send the File to this Email: ex(00:02), ex(17:00)");
                    string time = Console.ReadLine();

                    Console.WriteLine();
                    Console.WriteLine("Choose the Interval: ");
                    Console.WriteLine("1. in Days"
                        + "\n 2. in Hours"
                        + "\n 3. in Minutes");
                    operation = int.Parse(Console.ReadLine());
                    string interval = "";
                    if (operation == 1)
                    {
                        interval = "Days";
                    }
                    else if (operation == 2)
                    {
                        interval = "Hours";
                    }
                    else if (operation == 3)
                    {
                        interval = "Minutes";
                    }
                    Console.WriteLine("Enter the Repeat number: ");
                    double repeat = double.Parse(Console.ReadLine());

                    scheduledEmailToSend.Add(new Department(departmentName, HandlingLocalDb.getDepartmentEmail(departmentName, companyName), time, repeat, interval));
                }

                else if (operation == 2)
                {
                    //TODO validation for an empty department list
                    if (scheduledEmailToSend.Count == 0)
                    {
                        Console.WriteLine("You Should choose at least one department to continue");
                    }

                    // Go to QueryAndSendEmailInterface
                    else
                    {
                        AddScheduledQueryInterface(scheduleName, scheduledEmailToSend, companyName);
                        scheduledEmailToSend.Clear();
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
        /// Adding the Schedule task Queries
        /// </summary>
        /// <param name="scheduleName">Name of the Schedule Task</param>
        /// <param name="scheduledEmails">List of choosen departments Emails to send to</param>
        /// <param name="companyName">Name of the Company</param>
        public static void AddScheduledQueryInterface(string scheduleName, List<Department> scheduledEmails, string companyName)
        {
            List<Query> queryList = new List<Query>();

            string query;
            string sheetName;

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.Write("Enter your opertion number: ");
                Console.WriteLine("\n 1. Add sql query " +
                    "\n 2. Add the Excel File Information"
                    );
                operation = int.Parse(Console.ReadLine());
                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("Add your Query: ");
                    query = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Add Query Name: ");
                    string queryName = Console.ReadLine();
                    Console.WriteLine("Add the sheet name for this query: ");
                    sheetName = Console.ReadLine();
                    queryList.Add(new Query(query, sheetName, queryName));
                }

                else if (operation == 2)
                {
                    if (queryList.Count != 0)
                    {


                        Console.WriteLine();
                        Console.WriteLine("Enter the Excel file name (without extension): ");
                        string name = Console.ReadLine();

                        Console.WriteLine();
                        Console.WriteLine("Enter the file path: ");
                        string path = Console.ReadLine();

                        Console.WriteLine();
                        Console.WriteLine("Enter the Email \"Subject\": ");
                        string emailSubject = Console.ReadLine();

                        Console.WriteLine();
                        Console.WriteLine("Enter the Email \"Body\" text: ");
                        string emailBody = Console.ReadLine();

                        HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                        //TODO make function to add task in db
                        handlingLocalDb.AddScheduleTask(new ScheduleTask(scheduleName, emailSubject, emailBody, name, path, companyName, queryList, scheduledEmails));
                        queryList.Clear();
                        scheduledEmails.Clear();
                        break;
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("you must add a query first.");
                    }

                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("you must add a query first.");
                }
            }


        }

        /// <summary>
        /// Method Enables the User to Edit Schedule tasks and restart them after editing. 
        /// </summary>
        public static void EditScheduleTask()
        {
            Console.WriteLine();
            int choice;
            string taskName;
            Console.WriteLine("Choose Task: ");
            List<string> allTasks = HandlingLocalDb.ListOfAllTasks();
            for (int i = 0; i < allTasks.Count; i++)
            {
                Console.WriteLine((i + 1) + ". " + allTasks[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

            taskName = allTasks[choice - 1];
            allTasks.Clear();
            while (true)
            {
                Console.WriteLine("1. Edit Main Task (Task Name, Email Subject, Body, Excel file name and path)."
                + "\n 2. Edit Department send Time."
                + "\n 3. Edit Queries."
                + "\n 4. Add a new Department to this task"
                + "\n 5. Add a new Query to this Task"
                + "\n 6. Back to Settings");

                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    EditMainTaskInterface(taskName);
                    break;
                }
                else if (operation == 2)
                {
                    EditDepartmentSendTime(taskName);
                    break;
                }
                else if (operation == 3)
                {
                    EditQueries(taskName);
                    break;
                }
                else if (operation == 4)
                {
                    List<string> scheduledEmailToSend = HandlingLocalDb.ListOfAllDepartmentInATask(taskName);


                    List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentsNamesInACompany(HandlingLocalDb.getTaskCompanyName(taskName));


                    //TODO list of all company departments to choose from
                    if (scheduledEmailToSend.Count == allDepartments.Count)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("You've choosed all the departments");
                        Console.WriteLine();
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Choose the Departments");

                        for (int i = 0; i < allDepartments.Count; i++)
                        {
                            Console.WriteLine();
                            foreach (string department in scheduledEmailToSend)
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
                    string departmentName = allDepartments[choice - 1];
                    Console.WriteLine();

                    Console.WriteLine("Enter the Time you want to send the File to this Email: ex(00:02), ex(17:00)");
                    string time = Console.ReadLine();

                    Console.WriteLine();
                    Console.WriteLine("Choose the Interval: ");
                    Console.WriteLine("1. in Days"
                        + "\n 2. in Hours"
                        + "\n 3. in Minutes");
                    operation = int.Parse(Console.ReadLine());
                    string interval = "";
                    if (operation == 1)
                    {
                        interval = "Days";
                    }
                    else if (operation == 2)
                    {
                        interval = "Hours";
                    }
                    else if (operation == 3)
                    {
                        interval = "Minutes";
                    }
                    Console.WriteLine("Enter the Repeat number: ");
                    double repeat = double.Parse(Console.ReadLine());

                    HandlingLocalDb.AddSingleDepartmentToTask(new Department(departmentName, HandlingLocalDb.getDepartmentEmail(departmentName, HandlingLocalDb.getTaskCompanyName(taskName)), time, repeat, interval), taskName);
                    break;
                }
                else if (operation == 5)
                {
                    Console.WriteLine();
                    Console.WriteLine("Add your Query: ");
                    string query = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine("Add Query Name: ");
                    string Name = Console.ReadLine();
                    Console.WriteLine("Add the sheet name for this query: ");
                    string sheetName = Console.ReadLine();
                    HandlingLocalDb.AddSingleQueryToTask(new Query(query, sheetName, Name), taskName);
                    break;
                }

                else if (operation == 6)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a Wrong Input");


            }
        }


        /// <summary>
        /// Edit the Main Schedule Task Infrormation (Name, Email Data, Excel File Data)
        /// </summary>
        /// <param name="taskName">Name of the Choosen Schedule Task</param>
        public static void EditMainTaskInterface(string taskName)
        {
            while (true)
            {

                Console.WriteLine();
                Console.WriteLine("1. Edit Schedule Task Name"
                    + "\n 2. Edit Email Subject"
                    + "\n 3. Edit Email Body"
                    + "\n 4. Edit Excel File Name"
                    + "\n 5. Edit Excel File Path"
                    + "\n 6. Back to Edit Schedule Task");
                operation = int.Parse(Console.ReadLine());
                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Schedule Task Name is: " + taskName);
                    Console.WriteLine();
                    Console.WriteLine("Write a new Schedule Task Name: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateScheduleTaskName(taskName, newData);
                    taskName = newData;
                }
                else if (operation == 2)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Schedule Email Subject is: " + HandlingLocalDb.getTaskEmailSubject(taskName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Schedule Email Subject: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateScheduleEmailSubject(taskName, newData);
                }

                else if (operation == 3)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Schedule Email Body is: " + HandlingLocalDb.getTaskEmailBody(taskName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Schedule Email Body: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateScheduleEmailBody(taskName, newData);
                }
                else if (operation == 4)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Schedule Excel File Name is: " + HandlingLocalDb.getTaskExcelFileName(taskName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Schedule Excel File Name: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateScheduleExcelFileName(taskName, newData);
                }
                else if (operation == 5)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Schedule Excel File Path is: " + HandlingLocalDb.getTaskExcelFilePath(taskName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Schedule Excel File Path: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateScheduleExcelFilePath(taskName, newData);
                }
                else if (operation == 6)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a wrong Input.");
            }
        }

        /// <summary>
        /// Edit the choosen department Send Time.
        /// </summary>
        /// <param name="taskName">Name of the Schedule Task</param>
        public static void EditDepartmentSendTime(string taskName)
        {
            string departmentName;
            int choice;
            Console.WriteLine("Choose a Department to Edit");
            //TODO list of all companies in the data base to choose from
            List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentInATask(taskName);
            for (int i = 0; i < allDepartments.Count; i++)
            {
                Console.WriteLine((i + 1) + ". " + allDepartments[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());
            departmentName = allDepartments[choice - 1];
            allDepartments.Clear();
            Console.WriteLine();
            while (true)
            {
                string newData;
                Console.WriteLine();
                Console.WriteLine("1. Edit Send Time"
                    + "\n 2. Edit the Interval"
                    + "\n 3. Edit the Repeat"
                    + "\n 4. Back to Edit Schedule Task");
                operation = int.Parse(Console.ReadLine());
                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Send Time is: " + HandlingLocalDb.getEmailSendTime(taskName, departmentName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Send Time: ");
                    newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateSendEmailTime(taskName, departmentName, newData);
                }
                else if (operation == 2)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current send Interval is: " + HandlingLocalDb.getInterval(taskName, departmentName));
                    Console.WriteLine();

                    Console.WriteLine();
                    Console.WriteLine("Choose the Interval: ");
                    Console.WriteLine("1. in Days"
                        + "\n 2. in Hours"
                        + "\n 3. in Minutes");
                    operation = int.Parse(Console.ReadLine());
                    string interval = "";
                    if (operation == 1)
                    {
                        interval = "Days";
                    }
                    else if (operation == 2)
                    {
                        interval = "Hours";
                    }
                    else if (operation == 3)
                    {
                        interval = "Minutes";
                    }
                    Console.WriteLine("Enter the Repeat number: ");
                    double repeat = double.Parse(Console.ReadLine());
                    //TODO
                    HandlingLocalDb.UpdateInterval(taskName, departmentName, interval, repeat);
                }

                else if (operation == 3)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Repeat is: " + HandlingLocalDb.getRepeat(taskName, departmentName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Repeat: ");
                    double repeat = double.Parse(Console.ReadLine());
                    //TODO Update function
                    HandlingLocalDb.UpdateRepeat(taskName, departmentName, repeat);
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
        /// Edit the Schedule Task Queries Information.
        /// </summary>
        /// <param name="taskName">Name of the choosen Schedule Task</param>
        public static void EditQueries(string taskName)
        {
            string queryName;
            int choice;
            Console.WriteLine("Choose a Query to Edit");
            //TODO list of all companies in the data base to choose from
            List<string> allQueries = HandlingLocalDb.ListOfAllQueriesInTask(taskName);
            for (int i = 0; i < allQueries.Count; i++)
            {
                Console.WriteLine((i + 1) + ". " + allQueries[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());
            queryName = allQueries[choice - 1];
            allQueries.Clear();
            Console.WriteLine();
            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Edit Query Name"
                    + "\n 2. Edit Query"
                    + "\n 3. Edit Sheet Name"
                    + "\n 4. Back to Schedule Task");

                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Query Name is: " + HandlingLocalDb.getQueryName(taskName, queryName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Query Name: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateQueryName(taskName, queryName, newData);
                }

                else if (operation == 2)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Query is: " + HandlingLocalDb.getQuery(taskName, queryName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Query: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateQuery(taskName, queryName, newData);
                }
                else if (operation == 3)
                {
                    Console.WriteLine();
                    Console.WriteLine("the current Sheet Name is: " + HandlingLocalDb.getSheetName(taskName, queryName));
                    Console.WriteLine();
                    Console.WriteLine("Write a new Sheet Name: ");
                    string newData = Console.ReadLine();
                    //TODO Update function
                    HandlingLocalDb.UpdateSheetName(taskName, queryName, newData);
                }
                else if(operation == 4)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a Wrong Input");
            }
        }

        /// <summary>
        /// Deleting the Schedule Task Information.
        /// </summary>
        public static void DeleteScheduleTask()
        {
            Console.WriteLine();
            int choice;
            string taskName;
            Console.WriteLine("Choose Task: ");
            List<string> allTasks = HandlingLocalDb.ListOfAllTasks();
            for (int i = 0; i < allTasks.Count; i++)
            {
                Console.WriteLine((i + 1) + ". " + allTasks[i]);
            }

            Console.WriteLine();
            choice = int.Parse(Console.ReadLine());

            taskName = allTasks[choice - 1];
            allTasks.Clear();
            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("1. Delete the whole Task"
                    + "\n 2. Delete a Department from the Task"
                    + "\n 3. Delete a Query from the task"
                    + "\n 4. Back to Schedule Task Settings");
                operation = int.Parse(Console.ReadLine());

                if (operation == 1)
                {
                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " + taskName + "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        HandlingLocalDb.DeleteTask(taskName);

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
                    Console.WriteLine("Choose a Department to Delete");
                    //TODO list of all companies in the data base to choose from
                    List<string> allDepartments = HandlingLocalDb.ListOfAllDepartmentInATask(taskName);
                    for (int i = 0; i < allDepartments.Count; i++)
                    {
                        Console.WriteLine((i + 1) + ". " + allDepartments[i]);
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    departmentName = allDepartments[choice - 1];
                    allDepartments.Clear();
                    Console.WriteLine();

                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " + departmentName + "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        HandlingLocalDb.DeleteDepartmentFromTask(taskName, departmentName);
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
                    string queryName;
                    Console.WriteLine("Choose a Query to Delete");
                    //TODO list of all companies in the data base to choose from
                    List<string> allQueries = HandlingLocalDb.ListOfAllQueriesInTask(taskName);
                    for (int i = 0; i < allQueries.Count; i++)
                    {
                        Console.WriteLine((i + 1) + ". " + allQueries[i]);
                    }

                    Console.WriteLine();
                    choice = int.Parse(Console.ReadLine());
                    queryName = allQueries[choice - 1];
                    allQueries.Clear();
                    Console.WriteLine();

                    string agreement;
                    Console.WriteLine(" Are you sure you want to delet the " + queryName + "? (yes) or (no)");
                    agreement = Console.ReadLine().ToLower();
                    if (agreement == "yes")
                    {
                        HandlingLocalDb.DeleteQueryFromTask(taskName, queryName);
                        break;
                    }
                    else if (agreement == "no")
                    {
                        break;
                    }
                    else
                        Console.WriteLine("You've Entered a wrong input");
                }
                else if (operation == 4)
                {
                    break;
                }
                else
                    Console.WriteLine("You've Entered a Wrong Input");
            }
        }


    }
}
