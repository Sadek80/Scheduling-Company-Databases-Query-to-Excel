using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class Department
    {
        private string _name;
        private string _email;

        public Department(string name, string email)
        {
            this._name = name;
            this._email = email;
        }

        // Setters
        public void setName(string name)
        {
            this._name = name;
        }

        public void setEmail(string email)
        {
            this._email = email;
        }

        // Getters

        public string getName()
        {
            return this._name;
        }

        public string getEmail()
        {
            return this._email;
        }
    }
}
