using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class Query
    {
        private  string query;
        private  string querySheetName;
        public string queryName { get; set; }


        public Query(string query, string querySheetName)
        {
            this.query = query;
            this.querySheetName = querySheetName;
        }

        public Query(string query, string querySheetName, string queryName) : this(query, querySheetName)
        {
            this.queryName = queryName;
        }

        public string getQuery()
        {
            return query;
        }

        public string getQuerySheetName()
        {
            return querySheetName;
        }
    }
}
