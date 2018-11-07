using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics.Contracts;
using System.Text;
using System.Threading.Tasks;

namespace SearchEngine
{
    /// <summary>
    /// Performs search operations in the directory and links to a datasource 
    /// </summary>
    public class DirectoryOperations: FileHandler
    {
        private string directory = @"c:\";
        private static DatabaseEngine DBE;
        private static List<string> directoryDocuments;
        private static SortedDictionary<string, Keywords> index;

        /// <summary>
        /// Returns an instance of the DatabaseEngine
        /// </summary>
        public static DatabaseEngine DB
        {
            get { return DBE; }
        }

        /// <summary>
        /// Gets the searched items from the DatabaseEngine
        /// </summary>
        public List<string> SearchedTerms
        {
            get { return DBE.GetSearchedTerms(); }
        }

        /// <summary>
        /// Returns the index table
        /// </summary>
        private static SortedDictionary<String, Keywords> Index
        {
            get { return index; }
        }

        /// <summary>
        /// Gets all the documents in the directory
        /// </summary>
        public static List<string> DocumentsInDirectory
        {
            get { return directoryDocuments; }
        }

        /// <summary>
        /// The Directory to be searched
        /// </summary>
        /// <param name="directory">directory path name</param>
        public DirectoryOperations(string directory)
        {
            this.directory = directory;
            initializeDirectory();
            loadIndexTable();
        }

        /// <summary>
        /// Reads all the files in the given directory and writes all tokens and file paths to the database
        /// </summary>
        public void initializeDirectory()
        {

            directoryDocuments = new List<string>();
            DBE = new DatabaseEngine();
            List<string> storedDocuments = DBE.GetDocumentPaths();

            List<Dictionary<string, List<string>>> docsToBeStored = new List<Dictionary<string, List<string>>>();

            var files = Directory.GetFiles(directory);
            foreach (var file in files)
            {
                if (!storedDocuments.Contains(file))
                {
                    docsToBeStored.Add(readFiles(file));
                }
                directoryDocuments.Add(file);

            }
            writeToDatabase(docsToBeStored);
        }

        /// <summary>
        /// Re-initializes the directory, checks if new files have been added 
        /// and updates the database if so, and gets the index table
        /// </summary>
        public void refreshDirectory()
        {
            initializeDirectory();
            foreach (string file in directoryDocuments)
            {
                Dictionary<string, List<string>> dict = readFiles(file);
                foreach (KeyValuePair<string, List<string>> k in dict)
                {
                    DBE.UpdateDocument(k.Key, k.Value);
                }
            }
            loadIndexTable();
        }

        /// <summary>
        /// Adds the path and lists of tokens in a document to the database
        /// </summary>
        /// <param name="docsToBeStored"></param>
        private void writeToDatabase(List<Dictionary<string, List<string>>> docsToBeStored)
        {
            foreach (Dictionary<string, List<string>> dict in docsToBeStored)
            {
                foreach (KeyValuePair<string, List<string>> k in dict)
                {
                    DBE.AddDocument(k.Key, k.Value);
                }
            }
        }

        /// <summary>
        /// Returns the index table which consists of the Word as the Key
        /// and an object that contains a list of all the files where that word can be found as the Value
        /// </summary>
        private void loadIndexTable()
        {
            index = WordDictionary.SingleWordDictionary(DBE.ReadDocuments(directoryDocuments));

        }

        /// <summary>
        /// This method searches for a given query, saves query and results
        /// Requires: The query value is not empty or null or whitespaces
        /// </summary>
        /// <exception cref="NullReferenceException">Thrown when the query string is null or consists of only white spaces</exception>
        /// <param name="query">the query string</param>
        /// <returns>an instance of the MatchResponseclass- A list of documents and the 
        /// time taken to retrieve those documents</returns>
        public MatchResponse Find(string query)
        {
            Contract.Requires<NullReferenceException>(!String.IsNullOrWhiteSpace(query));
            DBE.StoreSearchTerm(query);
            return WordDictionary.Match(index, query);
        }
    }
}
