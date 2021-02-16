using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class Department
    {
        private string _name;
        private string _email;
        private string Time;
        private int Repeat;
        private string Interval;

        public Department(string name, string email, string time, int repeat, string interval) : this(name, email)
        {
            Time = time;
            Repeat = repeat;
            Interval = interval;
        }

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

        public void setTime(string time)
        {
            Time = time;
        }

        public void setInterval(string Interval)
        {
            this.Interval = Interval;
        }

        public void setRepeat(int repeat)
        {
            this.Repeat = repeat;
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

        public string getTime()
        {
            return this.Time;
        }
        public string getInterval()
        {
            return this.Interval;
        }
        public int getRepeat()
        {
            return this.Repeat;
        }


    }
}
