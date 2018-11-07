using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics.Contracts;
using System.Threading.Tasks;

namespace SearchEngine
{
    /// <summary>
    /// Takes a document, extracts tokens and creates an index table showing words and the number of occurrences of each word
    /// </summary>
    public class WordDictionary
    {
        public static Dictionary<string, Document> docDataStore = new Dictionary<string, Document>();       //Key: Document Path, Value: Document Object
        private static string[] stopWords = { "a", "an", "the", "to", "be", "as", "at", "around",
                                            "are", "am", "always", "both", "do", "else", "if", "each", "from",
                                            "for", "to", "his", "here", "his", "so", "of", "off", "than", "that"
                                            ,"you", "your"};


        /// <summary>
        /// Documents Index Table generator
        /// Requires: The input to the index table generator is not null
        /// </summary>
        /// <exception cref="NullReferenceException">Thrown when the input to the index table is null</exception>
        /// <param name="listresult">
        ///  A dictionary. The filepaths are contained in the key, while the value contains the file wordlist
        /// </param>
        /// <returns>
        ///  A SortedDictionary where the Key is the Word and the Value is an
        ///  object that contains a list of all the files where that word can be found
        /// </returns>
        public static SortedDictionary<string, Keywords> SingleWordDictionary(Dictionary<string, List<string>> listresult)
        {
            Contract.Requires<NullReferenceException>(!listresult.Equals(null));
            Contract.Ensures(!Contract.Result<SortedDictionary<string, Keywords>>().Equals(null)); //ensures that the output index table is not null
            SortedDictionary<string, Keywords> WordDict = new SortedDictionary<string, Keywords>();
            Dictionary<string, Document> docData = new Dictionary<string, Document>();


            foreach (KeyValuePair<string, List<String>> x in listresult)
            {
                for (int i = 0; i < listresult[x.Key].Count; i++)
                {
                    if (!stopWords.Contains(listresult[x.Key][i]))
                    {
                        //word not in dictionary
                        if (!WordDict.ContainsKey(listresult[x.Key][i]))
                        {
                            Keywords t = new Keywords(listresult[x.Key][i]);
                            Document d = new Document(x.Key, i);
                            if (!docData.ContainsKey(d.DocID))
                            {
                                docData.Add(d.DocID, d);
                            }
                            t.AddAffiliatedDocument(d);
                            WordDict.Add(listresult[x.Key][i], t);
                        }
                        //word already in dictionary
                        else
                        {
                            Document d = new Document(x.Key, i);
                            if (!docData.ContainsKey(d.DocID))
                            {
                                docData.Add(d.DocID, d);
                            }
                            WordDict[listresult[x.Key][i]].AddAffiliatedDocument(d);
                        }
                    }
                }
            }
            WordDict.OrderBy(key => key.Key);
            docDataStore = docData;

            return WordDict;
        }


        /// <summary>
        /// Takes the users query string and returns a list of files that have the keywords
        /// of the search in it, sorted by their relevance to the query in inverse order
        /// 
        /// Requires: The query value is not empty or null or whitespaces
        /// </summary>
        /// <exception cref="NullReferenceException">Thrown when an empty or null query is entered</exception>
        /// <param name="Indexes">index table</param>
        /// <param name="query">the string the user enters</param>
        /// <returns>Files keyword list, sorted in inverse order according to relevance to query, and the time taken</returns>
        public static MatchResponse Match(SortedDictionary<string, Keywords> Indexes, string query)
        {
            Contract.Requires<NullReferenceException>(!String.IsNullOrWhiteSpace(query));
                        
            char[] ch = { ' ', '.', '-', ',', ';', '+', '!', ':', '@', '#', '^', '*' };
            List<string> wordsInQuery = query.ToLower().Split(ch).ToList();
            List<Document> relevantDocs = new List<Document>();
            Dictionary<string, float> termRelevance = new Dictionary<string, float>();
            SortedDictionary<string, float> Index = new SortedDictionary<string, float>();

            DateTime startTime = Convert.ToDateTime(DateTime.Now.ToLongTimeString());
            //get word relevance in each document  
            foreach (string word in wordsInQuery)
            {
                if (Indexes.ContainsKey(word))
                {
                    float Wordrelevance = 0;
                    foreach (Document doc in Indexes[word].Documents)
                    {
                        if (!relevantDocs.Contains(Indexes[word].Documents.Find(n => n.DocID == docDataStore[doc.DocID].DocID)))
                        {
                            Wordrelevance += WeighTerm(word, doc, Indexes);
                            relevantDocs.Add(doc);
                        }
                    }
                    termRelevance[word] = Wordrelevance;
                }
            }

            //get documents ranking
            foreach (Document doc in relevantDocs)
            {
                Index[doc.DocID] = DocRankValue(doc, wordsInQuery, Indexes, termRelevance);
            }

            List<KeyValuePair<string, float>> myList = Index.ToList();
            myList.Sort((firstPair, nextPair) => { return nextPair.Value.CompareTo(firstPair.Value); });
            MatchResponse response = new MatchResponse();
            List<string> topDocuments = myList.ToDictionary(key => { return key.Key.ToLower(); }).Keys.ToList();
            DateTime stopTime = Convert.ToDateTime(DateTime.Now.ToLongTimeString());       //stop
            response.TopDocuments = topDocuments;
            response.ResponseTime = (stopTime - startTime).TotalSeconds;
            return response;

        }

        /// <summary>
        /// Gets the rank value for a document on a query
        /// </summary>
        /// <param name="d">the current document</param>
        /// <param name="terms">the query in a tokenized form</param>
        /// <param name="Indexes">the index table</param>
        /// <param name="termRelevance">A dictionary which holds all the terms and their relevance factors</param>
        /// <returns>the rank value for a document on a query</returns>
        private static float DocRankValue(Document d, List<string> terms, SortedDictionary<string, Keywords> Indexes, Dictionary<string, float> termRelevance)
        {
            float documentRelevance = 0;
            foreach (string singleterm in terms)
            {
                try
                {
                    if (Indexes[singleterm].Documents.Contains(Indexes[singleterm].Documents.Find(n => n.DocID == d.DocID)))
                    {
                        documentRelevance += termRelevance[singleterm];
                    }
                }
                catch (KeyNotFoundException) { }
            }
            return documentRelevance;
        }

        /// <summary>
        /// Weighs a term in a document according to its relevance 
        /// </summary>
        /// <param name="term">the term</param>
        /// <param name="doc">the document</param>
        /// <param name="Indexes">the index table</param>
        /// <returns>the relevance factor for a term/token in a document</returns>
        private static float WeighTerm(string term, Document doc, SortedDictionary<String, Keywords> Indexes)
        {
            int TermFrequency = Indexes[term].Documents.FindAll(n => n.DocID == doc.DocID).Count;
            float InverseDocumentFrequency = (float)Math.Log10(DirectoryOperations.DocumentsInDirectory.Count / Indexes[term].NumberOfDocsWithTerm());
            float tfIdf = TermFrequency * InverseDocumentFrequency;
            return tfIdf;
        }




        //Next steps
        //1. Two-word dictionary entries. 

        //2. Rank Searches. Hint
        ////singlewords
        //foreach FileDictionary[docmuentname]
        // singledict.add[FileDictionary[docmuentname]]
        // singledict.sort

        ////doublelewords
        //foreach FileDictionary[docmuentname]
        // doubledict.add[FileDictionary[docmuentname]].;;.
        //doubledict.sort

        ////do scoring using tf-idf
    }

}
