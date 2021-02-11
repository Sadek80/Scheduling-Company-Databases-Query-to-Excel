using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    class Query
    {
        private  string query;
        private  string querySheetName;


        public Query(string query, string querySheetName)
        {
            this.query = query;
            this.querySheetName = querySheetName;
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
