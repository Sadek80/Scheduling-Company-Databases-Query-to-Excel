using System;
using System.Collections.Generic;
using System.Text;

namespace test_all_features_2
{
    /// <summary>
    /// Class for Organizing the Query Data so it can be easy to control the query information.
    /// </summary>
    class Query
    {
        /// <summary>
        /// Query Statement
        /// </summary>
        private string query;

        /// <summary>
        /// Sheet Name of a Query
        /// </summary>
        private string querySheetName;

        /// <summary>
        /// Name of a Query, it's used for Schedule Task Data Base.
        /// </summary>
        public string queryName { get; set; }

        /// <summary>
        /// Public Constructor For setting up the default Query Data
        /// </summary>
        /// <param name="query">Query Statement</param>
        /// <param name="querySheetName">Name of a Query, it's used for Schedule Task Data Base.</param>
        public Query(string query, string querySheetName)
        {
            this.query = query;
            this.querySheetName = querySheetName;
        }

        /// <summary>
        /// Public Constructor For setting up  the Query Data for the Schedule Tasks
        /// </summary>
        /// <param name="query"></param>
        /// <param name="querySheetName"></param>
        /// <param name="queryName"></param>
        public Query(string query, string querySheetName, string queryName) : this(query, querySheetName)
        {
            this.queryName = queryName;
        }


        /// <summary>
        /// Get the Query Statement
        /// </summary>
        /// <returns>Query Statement</returns>
        public string getQuery()
        {
            return query;
        }

        /// <summary>
        /// Get the Sheet Name of a Query
        /// </summary>
        /// <returns>Sheet Name</returns>
        public string getQuerySheetName()
        {
            return querySheetName;
        }
    }
}
