using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Organizing the Schedule Task Data
    /// </summary>
    class ScheduleTask
    {
        /// <summary>
        /// Schedule Name Propety
        /// </summary>
        public string ScheduleName { get; set; }

        /// <summary>
        /// Email Subject Shared across the Departments selected in the Schedule Task
        /// </summary>
        public string EmailSubject { get; set; }

        /// <summary>
        /// Email Body Shared across the Departments selected in the Schedule Task
        /// </summary>
        public string EmailBody { get; set; }

        /// <summary>
        /// Excel File Starting Name
        /// </summary>
        public string ExcelFileName { get; set; }

        /// <summary>
        /// Excel File Directory Path
        /// </summary>
        public string ExcelFilePath { get; set; }
        public string CompanyName { get; set; }

        /// <summary>
        /// List of Queries Shared across the Departments selected in the Schedule Task
        /// </summary>
        private List<Query> Queries;

        /// <summary>
        /// List of Selected Departments in the Schedule Task
        /// </summary>
        private List<Department> Emails;


        public ScheduleTask()
        {

        }

        public ScheduleTask(string scheduleName, string emailSubject, string emailBody, string excelFileName, string excelFilePath, string companyName, List<Query> queries, List<Department> emails)
        {
            ScheduleName = scheduleName;
            EmailSubject = emailSubject;
            EmailBody = emailBody;
            ExcelFileName = excelFileName;
            ExcelFilePath = excelFilePath;
            CompanyName = companyName;
            Queries = queries;
            Emails = emails;
        }

        public void setQueries(List<Query> queries)
        {
            Queries = queries;
        }

        public void setEmails(List<Department> emails)
        {
            Emails = emails;
        }

        public List<Query> getQueries()
        {
            return Queries;
        }

        public List<Department> getEmails()
        {
            return Emails;
        }

       
    }
}
