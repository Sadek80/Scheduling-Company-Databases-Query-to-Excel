using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Handling the local Data Base
    /// for Companies, Departments and Schedule Tasks Data
    /// </summary>
    class HandlingLocalDb
    {

        // Setting up the Connection String
        private static string path = System.IO.Path.GetFullPath(Environment.CurrentDirectory);
        private static string dataBaseName = "Database1.mdf";
        private static string ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + path + @"\" + dataBaseName + ";Integrated Security=True";

        // Make SQL Connections
        private static SqlConnection conection = new SqlConnection(ConnectionString);
        private static SqlConnection conection2 = new SqlConnection(ConnectionString);
        private static SqlConnection conection3 = new SqlConnection(ConnectionString);
        private static SqlConnection conection4 = new SqlConnection(ConnectionString);
        private static SqlConnection conection5 = new SqlConnection(ConnectionString);


        // Make SQL Commands for Parallel Usage
        private static SqlCommand command;
        private static SqlCommand command2;
        private static SqlCommand command3;
        private static SqlCommand command4;
        private static SqlCommand command5;



        ////////////////////////////////////
        
        //     Adding to the Data Base

        ////////////////////////////////////

        /// <summary>
        /// Add a company to the DataBase
        /// </summary>
        /// <param name="company">the company that user want to add</param>
        public static void AddCompany(Company company)
        {
            try
            {

                // Add the company name to the Companies table
                command = new SqlCommand("INSERT INTO [Companies] VALUES(@CompanyName)", conection);
                command.CommandType = CommandType.Text;
                command.Parameters.AddWithValue("@CompanyName", company.getCompanyName());
                conection.Open();
                command.ExecuteNonQuery();

                // Add the company departments to the Departments table
                foreach (Department department in company.getCompanyDepartments())
                {
                    command = new SqlCommand("INSERT INTO [Departments] (CompanyName, DepartmentName, DepartmentEmail) VALUES(@CompanyName, @DepartmentName, @DepartmentEmail)", conection);
                    command.CommandType = CommandType.Text;
                    command.Parameters.AddWithValue("@CompanyName", company.getCompanyName());
                    command.Parameters.AddWithValue("@DepartmentName", department.getName());
                    command.Parameters.AddWithValue("@DepartmentEmail", department.getEmail());
                    command.ExecuteNonQuery();
                }
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Company was added successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Add Company: " + e.Message);
                Console.ResetColor();

            }

        }

        /// <summary>
        /// Add a single Department
        /// </summary>
        /// <param name="companyName">The Company name of the department</param>
        /// <param name="department">The Department name</param>
        public static void AddDepartment(string companyName, Department department)
        {
            try
            {
                command = new SqlCommand("INSERT INTO [Departments] (CompanyName, DepartmentName, DepartmentEmail) VALUES(@CompanyName, @DepartmentName, @DepartmentEmail)", conection);
                command.CommandType = CommandType.Text;
                command.Parameters.AddWithValue("@CompanyName", companyName);
                command.Parameters.AddWithValue("@DepartmentName", department.getName());
                command.Parameters.AddWithValue("@DepartmentEmail", department.getEmail());
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department was added successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Add Department: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Add a Schedule Task
        /// </summary>
        /// <param name="task"></param>
        public void AddScheduleTask(ScheduleTask task)
        {
            try
            {
                // Add the Scheduled Task to the Schedule_Task table
                command = new SqlCommand("INSERT INTO [Schedule_Task] (ScheduleName, EmailSubject, EmailBody, ExcelFileName, ExcelFilePath, CompanyName)" +
                    " VALUES(@ScheduleName, @EmailSubject, @EmailBody, @ExcelFileName, @ExcelFilePath, @CompanyName)", conection);
                command.CommandType = CommandType.Text;
                command.Parameters.AddWithValue("@ScheduleName", task.ScheduleName);
                command.Parameters.AddWithValue("@EmailSubject", task.EmailSubject);
                command.Parameters.AddWithValue("@EmailBody", task.EmailBody);
                command.Parameters.AddWithValue("@ExcelFileName", task.ExcelFileName);
                command.Parameters.AddWithValue("@ExcelFilePath", task.ExcelFilePath);
                command.Parameters.AddWithValue("@CompanyName", task.CompanyName);


                conection.Open();
                command.ExecuteNonQuery();

                // Add the Scheduled Emails to the Schedule_Emails table
                if (task.getEmails().Count == 0)
                {
                    Console.WriteLine("error in data");
                }
                foreach (Department scheduledEmail in task.getEmails())
                {
                    command = new SqlCommand("INSERT INTO [Schedule_Emails] (ScheduleName, Email, Time, Interval, Repeat, DepartmentName)" +
                        "VALUES(@ScheduleName, @Email, @Time, @Interval, @Repeat, @DepartmentName)", conection);
                    command.CommandType = CommandType.Text;
                    command.Parameters.AddWithValue("@ScheduleName", task.ScheduleName);
                    command.Parameters.AddWithValue("@Email", scheduledEmail.getEmail());
                    command.Parameters.AddWithValue("@Time", scheduledEmail.getTime());
                    command.Parameters.AddWithValue("@Interval", scheduledEmail.getInterval());
                    command.Parameters.AddWithValue("@Repeat", scheduledEmail.getRepeat());
                    command.Parameters.AddWithValue("@DepartmentName", scheduledEmail.getName());


                    command.ExecuteNonQuery();
                }

                // Add the Queries List to the Schedule_Query table
                foreach (Query query in task.getQueries())
                {
                    command = new SqlCommand("INSERT INTO [Schedule_Query] (ScheduleName, Query, SheetName, QueryName)" +
                        " VALUES(@ScheduleName, @Query, @SheetName, @QueryName)", conection);
                    command.CommandType = CommandType.Text;
                    command.Parameters.AddWithValue("@ScheduleName", task.ScheduleName);
                    command.Parameters.AddWithValue("@Query", query.getQuery());
                    command.Parameters.AddWithValue("@SheetName", query.getQuerySheetName());
                    command.Parameters.AddWithValue("@QueryName", query.queryName);


                    command.ExecuteNonQuery();
                }
                conection.Close();



                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Task was added successfuly!");
                Console.ResetColor();
                
                // Starting Running the Task in the Background After Adding it
                runTask(task);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Add Schedule Task: " + e.Message);
                Console.ResetColor();

            }

        }

        /// <summary>
        /// Add more Department to the Schedule Task
        /// </summary>
        /// <param name="department">Department Data</param>
        /// <param name="taskName">Schedule Task Name</param>
        public static void AddSingleDepartmentToTask(Department department, string taskName)
        {
            try
            {
                command = new SqlCommand("INSERT INTO [Schedule_Emails] (ScheduleName, Email, Time, Interval, Repeat, DepartmentName)" +
                            "VALUES(@ScheduleName, @Email, @Time, @Interval, @Repeat, @DepartmentName)", conection);
                command.CommandType = CommandType.Text;
                command.Parameters.AddWithValue("@ScheduleName", taskName);
                command.Parameters.AddWithValue("@Email", department.getEmail());
                command.Parameters.AddWithValue("@Time", department.getTime());
                command.Parameters.AddWithValue("@Interval", department.getInterval());
                command.Parameters.AddWithValue("@Repeat", department.getRepeat());
                command.Parameters.AddWithValue("@DepartmentName", department.getName());

                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department was added to the task successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(taskName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Add a Department to a Task: " + e.Message);
                Console.ResetColor();

            }
        }



        ////////////////////////////////////

        //Reading from the local Data Base

        ////////////////////////////////////



        /// <summary>
        /// Get all the Companies names from the DataBase
        /// </summary>
        /// <returns>a list of all companies names</returns>
        public static List<string> ListOfAllCompaniesNames()
        {
            List<string> allCompaniesList = new List<string>();
            try
            {
                command = new SqlCommand("SELECT CompanyName FROM [Companies]", conection);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allCompaniesList.Add(reader["CompanyName"].ToString());
                }
                conection.Close();
                return allCompaniesList;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the List of company names: " + e.Message);
                return allCompaniesList;
            }
        }

        /// <summary>
        /// Get company departments names from DataBase
        /// </summary>
        /// <param name="CompanyName">Company Name</param>
        /// <returns>List of all Departments names in a company</returns>
        public static List<string> ListOfAllDepartmentsNamesInACompany(string CompanyName)
        {
            List<string> allDepartments = new List<string>();
            try
            {
                command = new SqlCommand("SELECT DepartmentName FROM [Departments] WHERE CompanyName = @name", conection);
                command.Parameters.AddWithValue("@name", CompanyName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allDepartments.Add(reader["DepartmentName"].ToString());
                }
                conection.Close();
                return allDepartments;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the Dapartments names list: " + e.Message);
                return allDepartments;
            }
        }

        /// <summary>
        /// Get all selected Departments Emails
        /// </summary>
        /// <param name="DepartmentName">Department Name</param>
        /// <returns>List of Emails</returns>
        public static List<string> ListOfAllDepartmentsEmails(string CompanyName)
        {
            List<string> allEmails = new List<string>();
            try
            {
                command = new SqlCommand("SELECT DepartmentEmail FROM [Departments] WHERE CompanyName = @name", conection);
                command.Parameters.AddWithValue("@name", CompanyName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allEmails.Add(reader["DepartmentEmail"].ToString());
                }
                conection.Close();
                return allEmails;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the Departments emails: " + e.Message);
                return allEmails;
            }
        }


        /// <summary>
        /// Get Dapartment Email by the its name and the company
        /// </summary>
        /// <param name="DepartmentName">Department Name</param>
        /// <param name="companyName">Company Name</param>
        /// <returns>Department Email</returns>
        public static string getDepartmentEmail(string DepartmentName, string companyName)
        {
            string email = "";
            try
            {
                command = new SqlCommand("SELECT DepartmentEmail FROM [Departments] WHERE DepartmentName = @name AND CompanyName = @companyName", conection);
                command.Parameters.AddWithValue("@name", DepartmentName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    email = reader["DepartmentEmail"].ToString();
                }
                conection.Close();
                return email;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return email;
            }
        }

        /// <summary>
        /// Get All the Schedule Tasks Names from the Local Data Base
        /// </summary>
        /// <returns>List of Tasks Names</returns>
        public static List<string> ListOfAllTasks()
        {

            List<string> allTasksList = new List<string>();
            try
            {
                command = new SqlCommand("SELECT ScheduleName FROM [Schedule_Task]", conection);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allTasksList.Add(reader["ScheduleName"].ToString());
                }
                conection.Close();
                return allTasksList;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the List of all Tasks names: " + e.Message);
                return allTasksList;
            }
        }


        /// <summary>
        /// List of All Departments Names Selected in a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>List of Departments Names</returns>
        public static List<string> ListOfAllDepartmentInATask(string taskName)
        {
            List<string> allDepartmentssList = new List<string>();
            try
            {
                command = new SqlCommand("SELECT DepartmentName FROM [Schedule_Emails] WHERE ScheduleName = @taskName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allDepartmentssList.Add(reader["DepartmentName"].ToString());
                }
                conection.Close();
                return allDepartmentssList;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the List of all Departments names in a Task: " + e.Message);
                return allDepartmentssList;
            }
        }


        /// <summary>
        /// List of All Queries Name Added in a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>List of Queries Names</returns>
        public static List<string> ListOfAllQueriesInTask(string taskName)
        {
            List<string> allQueriesList = new List<string>();
            try
            {
                command = new SqlCommand("SELECT QueryName FROM [Schedule_Query] WHERE ScheduleName = @taskName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allQueriesList.Add(reader["QueryName"].ToString());
                }
                conection.Close();
                return allQueriesList;
            }
            catch (Exception e)
            {
                Console.WriteLine("in the List of Queries names in a Task: " + e.Message);
                return allQueriesList;
            }
        }


        /// <summary>
        /// List of All Queries Added in a Schedule Tasks
        /// </summary>
        /// <param name="taskName">Schedule Name</param>
        /// <returns>List of Queries Objects</returns>
        public List<Query> ListOfAllQueriesOrderedInTime(string taskName)
        {
            List<Query> allQueriesList = new List<Query>();
            using (new SqlConnection(ConnectionString))
            {
                try
                {
                    command3 = new SqlCommand("SELECT Query, SheetName FROM [Schedule_Query] WHERE ScheduleName = @taskName", conection3);
                    command3.Parameters.AddWithValue("@taskName", taskName);

                    command3.CommandType = CommandType.Text;
                    conection3.Open();
                    SqlDataReader reader = command3.ExecuteReader();
                    while (reader.Read())
                    {
                        allQueriesList.Add(new Query(reader["Query"].ToString(), reader["SheetName"].ToString()));
                    }
                    conection3.Close();
                    return allQueriesList;
                }
                catch (Exception e)
                {
                    Console.WriteLine("in the List of Queries Objects in a task: " + e.Message);
                    return allQueriesList;
                }
            }

        }

        /// <summary>
        /// List of All Departments Selected in a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>List of Departments Objects</returns>
        public static List<Department> ListOfAllDepartmentsOrderedInTime(string taskName)
        {
            List<Department> allDepartmentList = new List<Department>();
            using (new SqlConnection(ConnectionString))
            {
                try
                {
                    command4 = new SqlCommand("SELECT DepartmentName, Email, Interval, Time, Repeat FROM [Schedule_Emails] WHERE ScheduleName = @taskName", conection4);
                    command4.Parameters.AddWithValue("@taskName", taskName);

                    command4.CommandType = CommandType.Text;
                    conection4.Open();
                    SqlDataReader reader = command4.ExecuteReader();
                    while (reader.Read())
                    {
                        allDepartmentList.Add(new Department(reader["DepartmentName"].ToString(), reader["Email"].ToString(), reader["Time"].ToString(), (Double)reader["Repeat"], reader["Interval"].ToString()));
                    }
                    conection4.Close();
                    return allDepartmentList;
                }
                catch (Exception e)
                {
                    Console.WriteLine("in the List of Departments Objects in a Task: " + e.Message);
                    return allDepartmentList;
                }
            }

        }


        /// <summary>
        /// Get The Email Subject Used in a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>Email Subject</returns>
        public static string getTaskEmailSubject(string taskName)
        {
            string subject = "";
            try
            {
                command = new SqlCommand("SELECT EmailSubject FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    subject = reader["EmailSubject"].ToString();
                }
                conection.Close();
                if (subject == "")
                {
                    Console.WriteLine("wrong subject");
                }
                return subject;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Task Email Subject: " + e.Message);
                Console.ResetColor();
                return subject;
            }
        }

        /// <summary>
        /// Get Email Body Used in a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Name</param>
        /// <returns>Email Body</returns>
        public static string getTaskEmailBody(string taskName)
        {
            string body = "";
            try
            {
                command = new SqlCommand("SELECT EmailBody FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    body = reader["EmailBody"].ToString();
                }
                conection.Close();
                return body;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Task Email Body: " + e.Message);
                Console.ResetColor();
                return body;
            }
        }

        /// <summary>
        /// Get Excel File Starting Name Generated By a Schedule Task
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>Excel File Starting Name</returns>
        public static string getTaskExcelFileName(string taskName)
        {
            string fileName = "";
            try
            {
                command = new SqlCommand("SELECT ExcelFileName FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    fileName = reader["ExcelFileName"].ToString();
                }
                conection.Close();
                return fileName;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Task Excel File Starting Name: " + e.Message);
                Console.ResetColor();
                return fileName;
            }
        }

        /// <summary>
        /// Get the Directory Path that the Schedule Task Generates the Excel Files in
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>Directory Path</returns>
        public static string getTaskExcelFilePath(string taskName)
        {
            string filePath = "";
            try
            {
                command = new SqlCommand("SELECT ExcelFilePath FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    filePath = reader["ExcelFilePath"].ToString();
                }
                conection.Close();
                return filePath;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Task Directoty Path: " + e.Message);
                Console.ResetColor();
                return filePath;
            }
        }

        /// <summary>
        /// Get Company Name where Schedule Task Send Data
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <returns>Company Name</returns>
        public static string getTaskCompanyName(string taskName)
        {
            string companyName = "";
            try
            {
                command = new SqlCommand("SELECT companyName FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    companyName = reader["companyName"].ToString();
                }
                conection.Close();
                return companyName;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get task company Name: " + e.Message);
                Console.ResetColor();
                return companyName;
            }
        }

        /// <summary>
        /// Get the Sending Email Time for a Department
        /// </summary>
        /// <param name="taskName">Schedule Task Name</param>
        /// <param name="departmentName">Departmetn Name</param>
        /// <returns>Sending Time</returns>
        public static string getEmailSendTime(string taskName, string departmentName)
        {
            string sendTime = "";
            try
            {
                command = new SqlCommand("SELECT Time FROM [Schedule_Emails] WHERE ScheduleName = @taskName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    sendTime = reader["Time"].ToString();
                }
                conection.Close();
                return sendTime;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Sending Email Time: " + e.Message);
                Console.ResetColor();
                return sendTime;
            }
        }

        /// <summary>
        /// Get the Task Time Period
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="departmentName"></param>
        /// <returns></returns>
        public static string getInterval(string taskName, string departmentName)
        {
            string interval = "";
            try
            {
                command = new SqlCommand("SELECT Interval FROM [Schedule_Emails] WHERE ScheduleName = @taskName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    interval = reader["Interval"].ToString();
                }
                conection.Close();
                return interval;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Sending Emaill Interval: " + e.Message);
                Console.ResetColor();
                return interval;
            }
        }

        /// <summary>
        /// Get the Duration Until Repeat
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="departmentName"></param>
        /// <returns></returns>
        public static string getRepeat(string taskName, string departmentName)
        {
            string repeat = "";
            try
            {
                command = new SqlCommand("SELECT Repeat FROM [Schedule_Emails] WHERE ScheduleName = @taskName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    repeat = reader["Repeat"].ToString();
                }
                conection.Close();
                return repeat;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Duration until Repeat: " + e.Message);
                Console.ResetColor();
                return repeat;
            }
        }

        /// <summary>
        /// Get Query name
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="queryName"></param>
        /// <returns></returns>
        public static string getQueryName(string taskName, string queryName)
        {
            string query = "";
            try
            {
                command = new SqlCommand("SELECT QueryName FROM [Schedule_Query] WHERE ScheduleName = @taskName AND QueryName = @queryName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@queryName", queryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    query = reader["QueryName"].ToString();
                }
                conection.Close();
                return query;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Query Name: " + e.Message);
                Console.ResetColor();
                return query;
            }
        }

        /// <summary>
        /// Get Query Statement
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="queryName"></param>
        /// <returns></returns>
        public static string getQuery(string taskName, string queryName)
        {
            string query = "";
            try
            {
                command = new SqlCommand("SELECT Query FROM [Schedule_Query] WHERE ScheduleName = @taskName AND QueryName = @queryName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@queryName", queryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    query = reader["Query"].ToString();
                }
                conection.Close();
                return query;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Query Statement: " + e.Message);
                Console.ResetColor();
                return query;
            }
        }

        /// <summary>
        /// Get Query Sheet Name
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="queryName"></param>
        /// <returns></returns>
        public static string getSheetName(string taskName, string queryName)
        {
            string sheetName = "";
            try
            {
                command = new SqlCommand("SELECT SheetName FROM [Schedule_Query] WHERE ScheduleName = @taskName AND QueryName = @queryName", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@queryName", queryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    sheetName = reader["SheetName"].ToString();
                }
                conection.Close();
                return sheetName;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Query Sheet Name: " + e.Message);
                Console.ResetColor();
                return sheetName;
            }
        }

        /// <summary>
        /// Get All the Tasks That Company Name Use
        /// </summary>
        /// <param name="companyName"></param>
        /// <returns></returns>

        public static List<ScheduleTask> ListOfAllTasksFromCompanyName(string companyName)
        {
            List<ScheduleTask> Tasks = new List<ScheduleTask>();
            try
            {
                command2 = new SqlCommand("SELECT ScheduleName, EmailSubject, EmailBody, ExcelFileName, ExcelFilePath FROM [Schedule_Task] WHERE CompanyName = @companyName", conection2);
                command2.Parameters.AddWithValue("@companyName", companyName);

                command2.CommandType = CommandType.Text;
                conection2.Open();
                SqlDataReader reader = command2.ExecuteReader();
                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                while (reader.Read())
                {
                    string taskName = reader["ScheduleName"].ToString();
                    Console.WriteLine(taskName);
                    Tasks.Add(new ScheduleTask(taskName, reader["EmailSubject"].ToString(), reader["EmailBody"].ToString(),
                        reader["ExcelFileName"].ToString(), reader["ExcelFilePath"].ToString(), companyName,
                        handlingLocalDb.ListOfAllQueriesOrderedInTime(taskName),
                        ListOfAllDepartmentsOrderedInTime(taskName)));
                }
                conection2.Close();
                return Tasks;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get all Tasks By Company Name: " + e.Message);
                Console.ResetColor();
                return Tasks;
            }
        }

        /// <summary>
        /// Get Task Object By its Name
        /// </summary>
        /// <param name="taskName"></param>
        /// <returns></returns>
        public static ScheduleTask getTaskByName(string taskName)
        {
            ScheduleTask task = new ScheduleTask();
            try
            {
                command5 = new SqlCommand("SELECT EmailSubject, EmailBody, ExcelFileName, ExcelFilePath, CompanyName FROM [Schedule_Task] WHERE ScheduleName = @taskName", conection5);
                command5.Parameters.AddWithValue("@taskName", taskName);

                command5.CommandType = CommandType.Text;
                conection5.Open();
                SqlDataReader reader = command5.ExecuteReader();
                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                while (reader.Read())
                {

                    task = (new ScheduleTask(taskName, reader["EmailSubject"].ToString(), reader["EmailBody"].ToString(),
                        reader["ExcelFileName"].ToString(), reader["ExcelFilePath"].ToString(), reader["CompanyName"].ToString(),
                        handlingLocalDb.ListOfAllQueriesOrderedInTime(taskName),
                        ListOfAllDepartmentsOrderedInTime(taskName)));
                }
                conection5.Close();
                return task;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Task By Name: " + e.Message);
                Console.ResetColor();
                return task;
            }
        }


        ////////////////////////////////////////
        
        //     Update the Local Data Base
        
        ////////////////////////////////////////


        /// <summary>
        /// Update the company name
        /// </summary>
        /// <param name="currentName">the current name of the company</param>
        /// <param name="newName">the new name of the company</param>
        public static void UpdateCompanyName(string currentName, string newName)
        {
            try
            {
                command = new SqlCommand("UPDATE Companies SET CompanyName = @newName WHERE CompanyName = @currentName", conection);
                command.Parameters.AddWithValue("@currentName", currentName);
                command.Parameters.AddWithValue("@newName", newName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Company Name was updated successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                foreach (ScheduleTask task in ListOfAllTasksFromCompanyName(newName))
                {
                    handlingLocalDb.AbortTaskThread(task);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("in the update company name: " + e.Message);
            }

        }

        /// <summary>
        /// Update the department name
        /// </summary>
        /// <param name="currentName">the current name of the department</param>
        /// <param name="newName">the new name of the department</param>
        public static void UpdateDepartmentName(string currentName, string newName, string companyName)
        {
            try
            {
                command = new SqlCommand("UPDATE Departments SET DepartmentName = @newName WHERE DepartmentName = @currentName AND CompanyName = @companyName", conection);
                command.Parameters.AddWithValue("@currentName", currentName);
                command.Parameters.AddWithValue("@newName", newName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();

                command = new SqlCommand("UPDATE Schedule_Emails SET DepartmentName = @newName WHERE DepartmentName = @currentName AND ScheduleName In (SELECT ScheduleName FROM Schedule_Task WHERE CompanyName = @companyName)", conection);
                command.Parameters.AddWithValue("@newName", newName);
                command.Parameters.AddWithValue("@currentName", currentName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.ExecuteNonQuery();


                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department Name was updated successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                foreach (ScheduleTask task in ListOfAllTasksFromCompanyName(companyName))
                {
                    handlingLocalDb.AbortTaskThread(task);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Department name: " + e.Message);
                Console.ResetColor();
            }

        }

        /// <summary>
        /// Update the Department Email
        /// </summary>
        /// <param name="departmentName"></param>
        /// <param name="newEmail"></param>
        /// <param name="companyName"></param>
        public static void UpdateDepartmentEmail(string departmentName, string newEmail, string companyName)
        {
            try
            {
                command = new SqlCommand("UPDATE Departments SET DepartmentEmail = @newEmail WHERE DepartmentName = @departmentName AND CompanyName = @companyName", conection);
                command.Parameters.AddWithValue("@newEmail", newEmail);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();

                command = new SqlCommand("UPDATE Schedule_Emails SET Email = @newEmail WHERE DepartmentName = @departmentName AND ScheduleName In (SELECT ScheduleName FROM Schedule_Task WHERE CompanyName = @companyName)", conection);
                command.Parameters.AddWithValue("@newEmail", newEmail);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.ExecuteNonQuery();

                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department Email was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                foreach (ScheduleTask task in ListOfAllTasksFromCompanyName(companyName))
                {
                    handlingLocalDb.AbortTaskThread(task);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Department email: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Schedule Task Name
        /// </summary>
        /// <param name="currentName"></param>
        /// <param name="newName"></param>
        public static void UpdateScheduleTaskName(string currentName, string newName)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Task SET ScheduleName = @newName WHERE ScheduleName = @currentName", conection);
                command.Parameters.AddWithValue("@newName", newName);
                command.Parameters.AddWithValue("@currentName", currentName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Task Name was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(newName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task Name: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Updat Schedule Task Email Subject
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="newSubject"></param>
        public static void UpdateScheduleEmailSubject(string scheduleName, string newSubject)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Task SET EmailSubject = @newSubject WHERE ScheduleName = @scheduleName", conection);
                command.Parameters.AddWithValue("@newSubject", newSubject);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Email Subject was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task Email Subject: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Schedule Task Email Body
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="newBody"></param>
        public static void UpdateScheduleEmailBody(string scheduleName, string newBody)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Task SET EmailBody = @newBody WHERE ScheduleName = @scheduleName", conection);
                command.Parameters.AddWithValue("@newBody", newBody);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Enail Body was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task Email Body: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update the Schedule Task Excel File Starting Name
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="newName"></param>
        public static void UpdateScheduleExcelFileName(string scheduleName, string newName)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Task SET ExcelFileName = @newName WHERE ScheduleName = @scheduleName", conection);
                command.Parameters.AddWithValue("@newName", newName);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Excel File Name was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task Excel File Name: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update the Excel File Directory
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="newPath"></param>
        public static void UpdateScheduleExcelFilePath(string scheduleName, string newPath)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Task SET ExcelFilePath = @newPath WHERE ScheduleName = @scheduleName", conection);
                command.Parameters.AddWithValue("@newPath", newPath);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Schedule Excel File Path was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task File Directory: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Schedule Task Email Sending Time for a Department 
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="departmentName"></param>
        /// <param name="newTime"></param>
        public static void UpdateSendEmailTime(string scheduleName, string departmentName, string newTime)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Emails SET Time = @newTime WHERE ScheduleName = @scheduleName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@newTime", newTime);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Send Time was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Sending Email Time: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Sending Email Period of Time
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="departmentName"></param>
        /// <param name="newInterval"></param>
        /// <param name="newRepeat"></param>
        public static void UpdateInterval(string scheduleName, string departmentName, string newInterval, double newRepeat)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Emails SET Interval = @newInterval, Repeat = @newRepeat  WHERE ScheduleName = @scheduleName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@newInterval", newInterval);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.Parameters.AddWithValue("@newRepeat", newRepeat);


                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Interval was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Interval: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update the Duration Until Repeat
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="departmentName"></param>
        /// <param name="newRepeat"></param>
        public static void UpdateRepeat(string scheduleName, string departmentName, double newRepeat)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Emails SET Repeat = @newRepeat WHERE ScheduleName = @scheduleName AND DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@newRepeat", newRepeat);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Repeat was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Repeat: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Query Name
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="QueryName"></param>
        /// <param name="newQueryName"></param>
        public static void UpdateQueryName(string scheduleName, string QueryName, string newQueryName)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Query SET QueryName = @newQueryName WHERE ScheduleName = @scheduleName AND QueryName = @QueryName", conection);
                command.Parameters.AddWithValue("@newQueryName", newQueryName);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@QueryName", QueryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Query Name was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Query Name: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update Schedule Task Query Statement
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="QueryName"></param>
        /// <param name="newQuery"></param>
        public static void UpdateQuery(string scheduleName, string QueryName, string newQuery)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Query SET Query = @newQuery WHERE ScheduleName = @scheduleName AND QueryName = @QueryName", conection);
                command.Parameters.AddWithValue("@newQuery", newQuery);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@QueryName", QueryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Query was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule Task Query Statement: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Update a Query Sheet Name
        /// </summary>
        /// <param name="scheduleName"></param>
        /// <param name="QueryName"></param>
        /// <param name="newSheetName"></param>
        public static void UpdateSheetName(string scheduleName, string QueryName, string newSheetName)
        {
            try
            {
                command = new SqlCommand("UPDATE Schedule_Query SET SheetName = @newSheetName WHERE ScheduleName = @scheduleName AND QueryName = @QueryName", conection);
                command.Parameters.AddWithValue("@newSheetName", newSheetName);
                command.Parameters.AddWithValue("@scheduleName", scheduleName);
                command.Parameters.AddWithValue("@QueryName", QueryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Sheet Name was Updated Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(scheduleName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Query Sheet Name: " + e.Message);
                Console.ResetColor();

            }
        }
        

        //////////////////////////////////////
        
        // Delete from The Local Data Base

        //////////////////////////////////////
        


        /// <summary>
        /// Delete a Company from Database
        /// </summary>
        /// <param name="companyName">the Name of the company that user want to delete</param>
        public static void DeleteCompany(string companyName)
        {
            HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
            foreach (ScheduleTask task in ListOfAllTasksFromCompanyName(companyName))
            {
                handlingLocalDb.AbortTaskThread(task.ScheduleName);
            }
            try
            {
                command = new SqlCommand("DELETE FROM Companies WHERE CompanyName = @companyName ", conection);
                command.Parameters.AddWithValue("@companyName", companyName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();

                conection.Close();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Company was Deleted Successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete company: " + e.Message);
                Console.ResetColor();
            }
        }


        /// <summary>
        /// Delete Department from Database
        /// </summary>
        /// <param name="departmentName">the Department name that user want to delete</param>
        public static void DeleteDepartment(string departmentName, string companyName)
        {
           

            try
            {
                command = new SqlCommand("DELETE FROM Departments WHERE DepartmentName = @departmentName AND CompanyName = @companyName ", conection);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.Parameters.AddWithValue("@companyName", companyName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();

                command = new SqlCommand("DELETE from Schedule_Emails WHERE DepartmentName = @departmentName AND ScheduleName In (SELECT ScheduleName FROM Schedule_Task WHERE CompanyName = @companyName)", conection);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.Parameters.AddWithValue("@companyName", companyName);
                command.ExecuteNonQuery();

                conection.Close();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department was Deleted Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();
                foreach (ScheduleTask task in ListOfAllTasksFromCompanyName(companyName))
                {
                    handlingLocalDb.AbortTaskThread(task);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;

                Console.WriteLine("in the delete Department: " + e.Message);
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Delete Schedule Task By Name
        /// </summary>
        /// <param name="taskName"></param>
        public static void DeleteTask(string taskName)
        {
            HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

            handlingLocalDb.AbortTaskThread(taskName);
            try
            {
                command = new SqlCommand("DELETE FROM Schedule_Task WHERE ScheduleName = @taskName ", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Task was Deleted Successfuly!");
                Console.ResetColor();

                
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete Task: " + e.Message);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Delete a Department from Schedule Task
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="departmentName"></param>
        public static void DeleteDepartmentFromTask(string taskName, string departmentName)
        {
            try
            {
                command = new SqlCommand("DELETE FROM Schedule_Emails WHERE ScheduleName = @taskName AND DepartmentName = @departmentName ", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@departmentName", departmentName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Department was Deleted Successfuly!");
                Console.ResetColor();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(taskName));
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete Department from Task: " + e.Message);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Delete a Query from Schedule Task
        /// </summary>
        /// <param name="taskName"></param>
        /// <param name="QueryName"></param>
        public static void DeleteQueryFromTask(string taskName, string QueryName)
        {
            try
            {
                command = new SqlCommand("DELETE FROM Schedule_Query WHERE ScheduleName = @taskName AND QueryName = @QueryName ", conection);
                command.Parameters.AddWithValue("@taskName", taskName);
                command.Parameters.AddWithValue("@QueryName", QueryName);

                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();

                HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                handlingLocalDb.AbortTaskThread(getTaskByName(taskName));

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Query was Deleted Successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete Query from Task: " + e.Message);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Run a Scheduled Task Thread
        /// </summary>
        /// <param name="task">Scheduled Task</param>
        public void runTask(ScheduleTask task)
        {

            lock (this)
            {
                try
                {
                    HandlingLocalDb handlingLocalDb = new HandlingLocalDb();

                    // Check if there are Departments and Queries in the Task to begin the Execution
                    if (handlingLocalDb.ListOfAllQueriesOrderedInTime(task.ScheduleName).Count != 0 && HandlingLocalDb.ListOfAllDepartmentsOrderedInTime(task.ScheduleName).Count != 0)
                    {
                       
                            List<Query> Queries = handlingLocalDb.ListOfAllQueriesOrderedInTime(task.ScheduleName);

                            // Run a new Thread for Every Department
                            foreach (Department department in HandlingLocalDb.ListOfAllDepartmentsOrderedInTime(task.ScheduleName))
                            {
                                int charLocation = department.getTime().IndexOf(':');
                                int Hour = int.Parse(department.getTime().Substring(0, charLocation));
                                int Minute = int.Parse(department.getTime().Substring(charLocation + 1));
                                Console.WriteLine(department.getName() + " Task is starting...");

                            Thread myThread = new Thread(() =>
                            {

                                int i = 1;

                                Scheduler.Instance.ScheduleTask(Hour, Minute, department.getInterval(), department.getRepeat(), task.ScheduleName + " - " + department.getName(), () =>
                                {

                                    Console.WriteLine(task.CompanyName + " - " + department.getName() + " File is creating...");

                                    // Lock this Block of Code so one Thread only can enter this Block at a time
                                    lock (this)
                                    {

                                        // Set the Excel file name and path
                                        HandlingExcelFile handlingExcelFile = new HandlingExcelFile(task.ExcelFileName + " - " + department.getName() + " " + i, task.ExcelFilePath);

                                        HandlingQuery handlingQuery = new HandlingQuery();

                                        // Executes all the queries
                                        foreach (Query query in Queries)
                                        {
                                            handlingQuery.excuteQuery(query);
                                        }

                                        // Make the Excel file
                                        handlingExcelFile.makeExcelFile(handlingQuery.getWorkSheetsList());


                                        // Send the Email
                                        SendingEmail sendingEmail = new SendingEmail(department.getEmail(), task.EmailSubject, task.EmailBody,
                                                    handlingExcelFile.getPath());


                                        sendingEmail.sendingEmail();
                                    }
                                    i++;
                                });

                            });
                            myThread.Name = Convert.ToString(task.ScheduleName + department.getName());

                            // Start the Thread
                            myThread.Start();

                            // the thread after will wait until this thread starts
                            myThread.Join();

                        }
                       
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("in the run task: " + e.Message);
                }
         
            }

        }

        /// <summary>
        /// Ending and restarting the Schedule Task
        /// </summary>
        /// <param name="task">Task Object</param>
        public void AbortTaskThread(ScheduleTask task)
        {

            foreach (Department department in ListOfAllDepartmentsOrderedInTime(task.ScheduleName))
            {
                foreach (string key in Scheduler.Instance.timerDictionary.Keys)
                {
                    if (key == task.ScheduleName + " - " + department.getName())
                    {
                        Scheduler.Instance.killTimer(key);
                    }
                }
            }

            // Restart
            runTask(task);
        }

        /// <summary>
        /// Ending the Schedule Task only
        /// </summary>
        /// <param name="ScheduleName"></param>
        public void AbortTaskThread(string ScheduleName)
        {
            foreach (Department department in ListOfAllDepartmentsOrderedInTime(ScheduleName))
            {
                foreach (string key in Scheduler.Instance.timerDictionary.Keys)
                {
                    if (key == ScheduleName + " - " + department.getName())
                    {
                        Scheduler.Instance.killTimer(key);
                    }
                }
            }
        }
     }

}
