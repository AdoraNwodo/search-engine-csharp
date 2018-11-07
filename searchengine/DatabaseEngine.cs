using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace SearchEngine
{
    /// <summary>
    /// Performs operations of saving, reading, retrieving and updating on the database
    /// </summary>
    public class DatabaseEngine
    {
        static string connString = Properties.Settings.Default.connStringSearchEngine;    //database connection string.
        OleDbConnection conn = new OleDbConnection(connString);
        

        /// <summary>
        /// Adds document information to the database.
        /// Document Information being the path of the file and the list of all keywords
        /// </summary>
        /// <param name="docPath">file path</param>
        /// <param name="tokens">keywords list</param>
        /// <returns>1 if successful and 0 otherwise</returns>
        public int AddDocument(string docPath, List<string> tokens)
        {
                        
                string query = @"INSERT INTO Document VALUES(@Path, @Tokens)";
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                cmd = conn.CreateCommand();
                cmd.CommandText = query;
                cmd.Parameters.AddWithValue("@Path", docPath);
                cmd.Parameters.AddWithValue("@Tokens", GenerateCommaSeparatedValuesFromList(tokens));
                int result = cmd.ExecuteNonQuery();
                conn.Close();              
                return result;
        }

        /// <summary>
        /// Reads keywords from documents in a file by specifying the file path.
        /// </summary>
        /// <param name="doc_paths">the file paths</param>
        /// <returns>
        /// A Dictionary having the file path as its key and the list of distinct tokens as its value
        /// </returns>
        public Dictionary<string, List<string>> ReadDocuments(List<string> doc_paths)
        {
            Dictionary<String, List<String>> result = new Dictionary<String, List<String>>();
            
                conn.Open();

                foreach (string path in doc_paths)
                {
                    DataTable dt = new DataTable();
                    dt.Clear();
                    string query = @"SELECT DISTINCT Tokens FROM Document WHERE Path = @Path";
                    OleDbDataAdapter da = new OleDbDataAdapter(query, conn);
                    da.SelectCommand.Parameters.AddWithValue("@Path", path);
                    da.Fill(dt);
                    DataTableReader dtr = new DataTableReader(dt);
                    while (dtr.Read())
                    {                        
                        List<string> tokens = GenerateListFromCsv(dtr.GetString(0));
                        result.Add(path, tokens);
                    }
                }
                conn.Close();
            return result;
        }
          
        /// <summary>
        /// Updates the record of a document in the database.
        /// </summary>
        /// <param name="docPath">the file path</param>
        /// <param name="tokens">list of keywords affiliated with the document</param>
        /// <returns>1 if successful and 0 otherwise</returns>
        public int UpdateDocument(string docPath, List<string> tokens)
        {
            
                conn.Open();
                string query = @"UPDATE Document SET Tokens = @Tokens WHERE Path = '" + docPath + "';";
                OleDbCommand command = new OleDbCommand();
                command = conn.CreateCommand();
                command.CommandText = query;
                command.Parameters.AddWithValue("@Tokens", GenerateCommaSeparatedValuesFromList(tokens));
                int result = command.ExecuteNonQuery();
                conn.Close();
            return result;
        }

        /// <summary>
        /// Returns the list of paths of the documents in the database
        /// </summary>
        /// <returns>the document paths</returns>
        public List<string> GetDocumentPaths()
        {
            List<string> paths = new List<string>();
            
                conn.Open();
                OleDbCommand command = new OleDbCommand();

                command = conn.CreateCommand();
                command.CommandText = "SELECT Path FROM Document";

                OleDbDataReader reader;
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    paths.Add(reader.GetString(0));
                }

                conn.Close();
            return paths;

        }

        /// <summary>
        /// Adds the terms that have been searched to a record
        /// </summary>
        /// <param name="term">the term to be stored</param>
        /// <returns>1 if successful and 0 otherwise.</returns>
        public int StoreSearchTerm(string term)
        {
                conn.Open();
                string query = @"INSERT INTO Searched_Terms VALUES ('" + term + "')";
                OleDbCommand command = new OleDbCommand();
                command = conn.CreateCommand();
                command.CommandText = query;
                int result = command.ExecuteNonQuery();
                conn.Close();
                return result;
                
        }

        /// <summary>
        /// Retrieves the searched items
        /// </summary>
        /// <returns>the searched terms</returns>
        public List<string> GetSearchedTerms()
        {
            List<string> terms = new List<string>();
            
                conn.Open();
                string query = @"SELECT Term FROM Searched_Terms";
                OleDbCommand command = new OleDbCommand();
                command = conn.CreateCommand();
                command.CommandText = query;
                OleDbDataReader reader;// = new OleDbDataReader();
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (!terms.Contains(reader.GetString(0)))
                    {
                        terms.Add(reader.GetString(0));
                    }
                }
                conn.Close();
            return terms;
        }

        
        /// <summary>
        /// Uses the list of strings to generate values that will be separated by commas
        /// </summary>
        /// <param name="listOfStrings">the list of strings</param>
        /// <returns>the string with commas seperating each value</returns>
       private string GenerateCommaSeparatedValuesFromList(List<string> listOfStrings)
        {
            StringBuilder strCsv = new StringBuilder();
           foreach (string item in listOfStrings)
           {
               strCsv.Append(item + ",");
           }

           // Return the csv
           return strCsv.ToString();
        }

        /// <summary>
        /// Generates a list of strings from the strings separated by commas
        /// </summary>
        /// <param name="csv">the comma seperated value string</param>
        /// <returns>the list of strings</returns>
        private List<string> GenerateListFromCsv(string csv)
       {
           List<string> listOfStrings = new List<string>();
           string[] csvArray = csv.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string item in csvArray)
            {
                listOfStrings.Add(item);
            }

            return listOfStrings;
       }

        
    }
}
