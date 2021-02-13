using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class Company
    {
        private string _companyName;
        private List<Department> _departments;

        public Company(string name, List<Department> departments)
        {
            this._companyName = name;
            this._departments = departments;
        }

        // Setters
        public void setCompanyName(string name)
        {
            this._companyName = name;
        }

        public void setCompanyDepartments(List<Department> departments)
        {
            this._departments = departments;
        }

        // Getters
        public string getCompanyName()
        {
            return this._companyName;
        }

        public List<Department> getCompanyDepartments()
        {
            return _departments;
        }

    }
}
