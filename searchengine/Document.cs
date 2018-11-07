using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchEngine
{
    /// <summary>
    /// Represents a Document 
    /// </summary>
    public class Document
    {
        private readonly string documentID;
        private readonly int PositionOfTerm;   

        /// <summary>
        /// Gets the documents ID
        /// </summary>
        public string DocID
        {
            get { return documentID; }
        }

        /// <summary>
        /// Gets the term's position in this Document 
        /// </summary>
        public int TermPositionInDocument
        {
            get { return PositionOfTerm; }
        }

        /// <summary>
        /// A new Document
        /// </summary>
        /// <param name="DocId">the document name</param>
        /// <param name="TermPosition">the term position in a document</param>
        public Document(string DocId, int TermPosition)
        {
            documentID = DocId;
            PositionOfTerm = TermPosition;
        }
    }
}
