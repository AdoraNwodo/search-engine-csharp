using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchEngine
{
    /// <summary>
    /// Contains the top documents after every search and the time used in searching
    /// </summary>
    public class MatchResponse
    {
        private static List<string> topDocuments;
        private static double responseTime;

        /// <summary>
        /// Represents the Relevant Documents after a search action
        /// </summary>
        public List<string> TopDocuments
        {
            get { return topDocuments; }
            set { topDocuments = value; }
        }

        /// <summary>
        /// Represents the time taken to complete a search action
        /// </summary>
        public double ResponseTime
        {
            get { return responseTime; }
            set { responseTime = value; }
        }
    }
}
