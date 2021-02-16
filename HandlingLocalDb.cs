using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace test_all_features_2
{
    class HandlingLocalDb
    {
        // Setting up the Connection String
        private static string path = System.IO.Path.GetFullPath(Environment.CurrentDirectory);
        private static string dataBaseName = "Database1.mdf";
        private static string ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + path + @"\" + dataBaseName + ";Integrated Security=True";

        // Make an SQL Connection
        private static SqlConnection conection = new SqlConnection(ConnectionString);
        private static SqlCommand command;


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
                Console.WriteLine("in the Handling Local Db: " + e.Message);
                Console.ResetColor();

            }

        }

        /// <summary>
        /// Add single Department
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
                Console.WriteLine("in the Handling Local Db: " + e.Message);
                Console.ResetColor();

            }
        }

        public static void AddScheduleTask(ScheduleTask task)
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
                if(task.getEmails().Count == 0)
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Handling Local Db: " + e.Message);
                Console.ResetColor();

            }

        }

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
        }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the Handling Local Db: " + e.Message);
                Console.ResetColor();

            }
        }
        // Read

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
        /// <param name="CompanyName"></param>
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
                Console.WriteLine("in the Dapartments names list: "+e.Message);
                return allDepartments;
            }
        }

        /// <summary>
        /// Get all selected Departments Emails
        /// </summary>
        /// <param name="DepartmentName"></param>
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
                Console.WriteLine("in the Departments emails: "+e.Message);
                return allEmails;
            }
        }


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
                Console.WriteLine("in the get Department email: "+e.Message);
                Console.ResetColor();
                return email;
            }
        }

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
                Console.WriteLine("in the List of company names: " + e.Message);
                return allTasksList;
            }
        }

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
                Console.WriteLine("in the List of company names: " + e.Message);
                return allDepartmentssList;
            }
        }

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
                Console.WriteLine("in the List of Queries: " + e.Message);
                return allQueriesList;
            }
        }

        public static string getTaskName(string taskName)
        {
            string name = "";
            try
            {
                command = new SqlCommand("SELECT ScheduleName FROM [Schedule_Task] WHERE ScheduleName = @name", conection);
                command.Parameters.AddWithValue("@name", taskName);
                command.CommandType = CommandType.Text;
                conection.Open();
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {

                    name = reader["ScheduleName"].ToString();
                }
                conection.Close();
                return name;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return name;
            }
        }

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
                if(subject == "")
                {
                    Console.WriteLine("wrong subject");
                }
                return subject;

            }

            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return subject;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return body;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return fileName;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return filePath;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return sendTime;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return interval;
            }
        }

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
                Console.WriteLine("in the get Department email: " + e.Message);
                Console.ResetColor();
                return repeat;
            }
        }

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
                Console.WriteLine("in the get Query: " + e.Message);
                Console.ResetColor();
                return query;
            }
        }

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
                Console.WriteLine("in the get Query: " + e.Message);
                Console.ResetColor();
                return query;
            }
        }

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
                Console.WriteLine("in the get Query: " + e.Message);
                Console.ResetColor();
                return sheetName;
            }
        }

        
        // Update

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
            }
            catch(Exception e)
            {
                Console.WriteLine("in the update company name: "+e.Message);
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Department name: " + e.Message);
                Console.ResetColor();           
            }

        }

        public static void UpdateDepartmentEmail(string departmentName , string newEmail, string companyName)
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Department email: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

        public static void UpdateInterval(string scheduleName, string departmentName, string newInterval, int newRepeat)
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

        public static void UpdateRepeat(string scheduleName, string departmentName, int newRepeat)
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Schedule: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Query: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Query: " + e.Message);
                Console.ResetColor();

            }
        }

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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Query: " + e.Message);
                Console.ResetColor();

            }
        }
        // Delete

        /// <summary>
        /// Delete a Company from Database
        /// </summary>
        /// <param name="companyName">the Name of the company that user want to delete</param>
        public static void DeleteCompany(string companyName)
        {
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
                Console.WriteLine("in the delete company: "+e.Message);
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;

                Console.WriteLine("in the delete company: " + e.Message);
                Console.ResetColor();

            }
        }

       public static void DeleteTask(string taskName)
        {
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
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete Task: " + e.Message);
                Console.ResetColor();
            }
        }

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

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Query was Deleted Successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the delete Task: " + e.Message);
                Console.ResetColor();
            }
        }
    }
}
