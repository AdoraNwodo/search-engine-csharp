using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchEngine
{
    /// <summary>
    /// The keyword present in a document
    /// </summary>
    public class Keywords
    {
        private readonly string TermName;                            
        private List<Document> ListOfDocsAffliatedWithTerm = new List<Document>();     //the documents where the term can be found in along with the position in each document

        /// <summary>
        ///  Term in a document
        /// </summary>
        /// <param name="nameOfTerm">the term name</param>
        public Keywords(String nameOfTerm)
        {
            TermName = nameOfTerm;
        }

        /// <summary>
        /// Gets the Term ID
        /// </summary>
        public string TermId
        {
            get { return TermName; }
        }

        /// <summary>
        /// returns the files that contain this term 
        /// </summary>
        public List<Document> Documents
        {
            get { return ListOfDocsAffliatedWithTerm; }
        }

        /// <summary>
        /// Adds a document that is affiliated with the term
        /// </summary>
        /// <param name="doc">the document</param>
        public void AddAffiliatedDocument(Document doc)
        {
            ListOfDocsAffliatedWithTerm.Add(doc);
        }

        /// <summary>
        /// Adds a document that is affiliated with the term
        /// </summary>
        /// <param name="docId">the document name</param>
        /// <param name="TermPositionInDocument">the position of the term in the document</param>
        public void AddAffiliatedDocument(string docId, int TermPositionInDocument)
        {
            Document d = new Document(docId, TermPositionInDocument);
            ListOfDocsAffliatedWithTerm.Add(d);
        }

        /// <summary>
        /// Computes the number of documents having a term
        /// </summary>
        /// <returns>the number of files having the term</returns>
        public int NumberOfDocsWithTerm()
        {
            List<Document> docList = new List<Document>();
            
            foreach (Document d in Documents)
            {
                Document files = Documents.Find(a => a.DocID == d.DocID);
                if(!docList.Contains(files))
                { docList.Add(d); }
            }
            return docList.Count;
        }

        

    }
}
