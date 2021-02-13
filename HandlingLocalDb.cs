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


        public static string getDepartmentEmail(string DepartmentName)
        {
            string email = "";
            try
            {
                command = new SqlCommand("SELECT DepartmentEmail FROM [Departments] WHERE DepartmentName = @name", conection);
                command.Parameters.AddWithValue("@name", DepartmentName);
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
        public static void UpdateDepartmentName(string currentName, string newName)
        {
            try
            {
                command = new SqlCommand("UPDATE Departments SET DepartmentName = @newName WHERE DepartmentName = @currentName", conection);
                command.Parameters.AddWithValue("@currentName", currentName);
                command.Parameters.AddWithValue("@newName", newName);
                command.CommandType = CommandType.Text;
                conection.Open();
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

        public static void UpdateDepartmentEmail(string departmentName , string newEmail)
        {
            try
            {
                command = new SqlCommand("UPDATE Departments SET DepartmentEmail = @newEmail WHERE DepartmentName = @departmentName", conection);
                command.Parameters.AddWithValue("@newEmail", newEmail);
                command.Parameters.AddWithValue("@departmentName", departmentName);
                command.CommandType = CommandType.Text;
                conection.Open();
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
        public static void DeleteDepartment(string departmentName)
        {
            try
            {
                command = new SqlCommand("DELETE FROM Departments WHERE DepartmentName = @departmentName ", conection);
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

                Console.WriteLine("in the delete company: " + e.Message);
                Console.ResetColor();

            }
        }

        public static void ChangeDefaultEmail(string email, string password)
        {
            try
            {
                command = new SqlCommand("UPDATE DeafaultEmail SET Email = @email, Password = @password WHERE Id = 1", conection);
                command.Parameters.AddWithValue("@email", email);
                command.Parameters.AddWithValue("@password", password);
                command.CommandType = CommandType.Text;
                conection.Open();
                command.ExecuteNonQuery();
                conection.Close();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Default Email was Updated Successfuly!");
                Console.ResetColor();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the update Department email: " + e.Message);
                Console.ResetColor();

            }
        }       
    }
}
