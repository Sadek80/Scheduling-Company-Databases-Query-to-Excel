using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class ScheduleTask
    {
        public string ScheduleName { get; set;}
        public string EmailSubject { get; set;}
        public string EmailBody { get; set; }
        public string ExcelFileName { get; set; }
        public string ExcelFilePath { get; set; }
        public string CompanyName { get; set; }
        private List<Query> Queries;
        private List<Department> Emails;

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
