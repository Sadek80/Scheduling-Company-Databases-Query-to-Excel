using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Organizing the Company Data
    /// </summary>
    class Company
    {
        /// <summary>
        /// Company Name
        /// </summary>
        private string _companyName;

        /// <summary>
        /// Company List of Departments
        /// </summary>
        private List<Department> _departments;


        /// <summary>
        /// public Constructor for Setting Up the Default Company Data
        /// </summary>
        /// <param name="name"></param>
        /// <param name="departments"></param>
        public Company(string name, List<Department> departments)
        {
            this._companyName = name;
            this._departments = departments;
        }


        /// <summary>
        /// Get the Company Name
        /// </summary>
        /// <returns>Company Naem</returns>
        public string getCompanyName()
        {
            return this._companyName;
        }


        /// <summary>
        /// Get the List of Compant Departments
        /// </summary>
        /// <returns>List of Company Departments</returns>
        public List<Department> getCompanyDepartments()
        {
            return _departments;
        }

    }
}
