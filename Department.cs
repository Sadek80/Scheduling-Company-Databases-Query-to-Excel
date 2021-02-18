using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Organizinf the Company Departments
    /// </summary>
    class Department
    {
        /// <summary>
        /// Department Name
        /// </summary>
        private string _name;

        /// <summary>
        /// Department Email 
        /// </summary>
        private string _email;

        /// <summary>
        /// Send Time 
        /// (24 Hours Clock Style)
        /// </summary>
        private string Time;

        /// <summary>
        /// Duration to Repeat the Email
        /// </summary>
        private double Repeat;

        /// <summary>
        /// Period of Time to send the Email
        /// (Minutes, Hours, Days)
        /// </summary>
        private string Interval;
        

        /// <summary>
        /// Public Constructor For setting up the Department data for the Schedule Task
        /// </summary>
        /// <param name="name">Department Name</param>
        /// <param name="email">Department Email</param>
        /// <param name="time">Send Time</param>
        /// <param name="repeat">Duration to Repeat the Email</param>
        /// <param name="interval">Period of Time</param>
        public Department(string name, string email, string time, double repeat, string interval) : this(name, email)
        {
            Time = time;
            Repeat = repeat;
            Interval = interval;
        }

        /// <summary>
        /// Public Constructor for Setting up the Default Department Data
        /// </summary>
        /// <param name="name">Department Name</param>
        /// <param name="email">Department Email</param>
        public Department(string name, string email)
        {
            this._name = name;
            this._email = email;
        }
      

        // Getters

        /// <summary>
        /// Get the Department Name
        /// </summary>
        /// <returns>Department Name</returns>
        public string getName()
        {
            return this._name;
        }


        /// <summary>
        /// Get the Department Email
        /// </summary>
        /// <returns>Deparment Email</returns>
        public string getEmail()
        {
            return this._email;
        }

        /// <summary>
        /// Get the Department Send Time
        /// </summary>
        /// <returns>Send Time</returns>
        public string getTime()
        {
            return this.Time;
        }

        /// <summary>
        /// Get Period of Time to send the Email
        /// </summary>
        /// <returns>Period of Time</returns>
        public string getInterval()
        {
            return this.Interval;
        }

        /// <summary>
        /// Get the Duration To repeat the Email
        /// </summary>
        /// <returns>Duration to Repeat</returns>
        public double getRepeat()
        {
            return this.Repeat;
        }

    }
}
