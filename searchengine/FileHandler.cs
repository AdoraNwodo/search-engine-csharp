using Code7248.word_reader;
using Excel;
using HtmlAgilityPack;
using Aspose.Slides;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Diagnostics.Contracts;
using Microsoft.Office.Core;
using Aspose.Slides.Util;
using System.Windows.Forms;
using System.Collections;



namespace SearchEngine
{
    /// <summary>
    /// Handles the reading and tokenizing of files with different formats
    /// </summary>
    public class FileHandler
    {
        /// <summary>
        /// Supported File extensions
        /// </summary>
        public static readonly ArrayList extensions = new ArrayList(){ ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".txt", ".html", ".xml" };

        /// <summary>
        /// Reads PDF file types and extracts the words in each PDF file
        /// Requires: The file path is in .pdf only
        /// </summary>
        /// <param name="filenameWithPath">path of PDF document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not a Portable Document
        /// Format file. 
        /// </exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        public static Dictionary<string, List<string>> readPdfFile(string filenameWithPath)
        {
            Contract.Requires<PlatformNotSupportedException>(System.IO.Path.GetExtension(filenameWithPath).Equals(".pdf"));
            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            PdfReader reader = new PdfReader(filenameWithPath);
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                ITextExtractionStrategy ITES = new LocationTextExtractionStrategy();
                string s = PdfTextExtractor.GetTextFromPage(reader, page, ITES);
                s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
                result.AddRange(s.Trim().ToLower().Split(new string[] { "\t\n\r", " " }, StringSplitOptions.RemoveEmptyEntries));
            }

            listresult.Add(filenameWithPath, result);

            return listresult;
        }

        /// <summary>
        /// Reads DOC and DOCX file types and extracts the words in each file
        /// Requires: The file path is in doc or docx format only
        /// </summary>
        /// <param name="filenameWithPath">path of DOC or DOCX document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of supported
        /// doc format. 
        /// </exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        internal static Dictionary<string, List<string>> readDocFiles(string filenameWithPath)
        {
            Contract.Requires<PlatformNotSupportedException>(System.IO.Path.GetExtension(filenameWithPath).Equals(".doc") ||
                System.IO.Path.GetExtension(filenameWithPath).Equals(".docx"));
            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            TextExtractor extractor = new TextExtractor(filenameWithPath);
            string temp = extractor.ExtractText().Trim();

            result.AddRange(temp.Split(new string[] { "\t\r\n", " " }, StringSplitOptions.RemoveEmptyEntries));

            listresult.Add(filenameWithPath, result);
            return listresult;
        }

        /// <summary>
        /// Reads PPT and PPTX file types and extracts the words in each file
        /// Requires: The file path is in ppt or pptx format only
        /// </summary>
        /// <param name="filenameWithPath">path of PPT and PPTX document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of supported
        /// presentation format. </exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        private static Dictionary<string, List<string>> readPPTFiles(string filenameWithPath)
        {
            Contract.Requires<PlatformNotSupportedException>(System.IO.Path.GetExtension(filenameWithPath).Equals(".ppt") ||
                System.IO.Path.GetExtension(filenameWithPath).Equals(".pptx"));

            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(filenameWithPath);
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                foreach (var item in presentation.Slides[i + 1].Shapes)
                {
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            string text = textRange.Text.ToLower().Trim().ToString();
                            result.AddRange(text.Split(new char[] {' ','\n','\t', '\r'}));
                        }
                    }
                }
            }
            PowerPoint_App.Quit();
            listresult.Add(filenameWithPath, result);
            return listresult;
        }

        /// <summary>
        /// Reads XLS and XLSX file types and extracts the words in each file.
        /// Requires: The file path is in xls or xlsx format only
        /// </summary>
        /// <param name="filenameWithPath">path of XLS and XLSX document including filename</param>
        /// <param name="extension">the excel file extension. xls or xlsx</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of supported
        /// spreadsheet format. 
        /// </exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        internal static Dictionary<string, List<string>> readExcelFiles(string filenameWithPath, string extension)
        {
            Contract.Requires<PlatformNotSupportedException>(extension.Equals(".xls") || extension.Equals(".xlsx"));
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            List<string> result = new List<string>();
            FileStream stream = File.Open(filenameWithPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelDataReader;

            if (extension.Replace(".","").Equals("xls"))
                excelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else
                excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            excelDataReader.IsFirstRowAsColumnNames = false;
            DataSet Workbook = excelDataReader.AsDataSet();
            StringBuilder words = new StringBuilder();
            for (int j = 0; j < Workbook.Tables.Count; j++)
            {
                DataTable Worksheet = Workbook.Tables[j];

                IEnumerable<DataRow> dt = from DataRow row in Worksheet.Rows select row;
                foreach (DataRow datarow in dt)
                {
                    if (datarow == null)
                    {
                        break;
                    }
                    for (int k = 0; k < Worksheet.Columns.Count; k++)
                    {
                        string[] strarr = datarow[k].ToString().Split();
                        foreach (string str in strarr)
                        {
                            words.Append(str + " ");
                        }
                    }
                }
            }
            result.AddRange(words.ToString().ToLower().Split(new string[] { "\r\n", " " }, StringSplitOptions.RemoveEmptyEntries));

            
            listresult.Add(filenameWithPath, result);
            return listresult;
        }

        /// <summary>
        /// Reads TXT file types and extracts the words in each TXT file
        /// Requires: The file path is in txt format only
        /// </summary>
        /// <param name="filenameWithPath">path of TXT document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of supported
        /// text format. 
        /// </exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        private static Dictionary<string, List<string>> readTxtFiles(string filenameWithPath)
        {
            Contract.Requires<PlatformNotSupportedException>(System.IO.Path.GetExtension(filenameWithPath).Equals(".txt"));

            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            StreamReader sR = File.OpenText(filenameWithPath);
            string temp = sR.ReadToEnd().Trim();
            result.AddRange(temp.Split(new string[] { "\t\r\n", " " }, StringSplitOptions.RemoveEmptyEntries));

            listresult.Add(filenameWithPath, result);
            return listresult;

        }

        /// <summary>
        /// Reads XML file types and extracts the words in each XML file
        /// Requires: The file path is in xml format only
        /// </summary>
        /// <param name="filenameWithPath">path of XML document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of 
        /// xml format.</exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        private static Dictionary<string, List<string>> readXMLFiles(string filenameWithPath)
        {
            Contract.Requires<PlatformNotSupportedException>(System.IO.Path.GetExtension(filenameWithPath).Equals(".xml"));
            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            XmlTextReader reader = new XmlTextReader(filenameWithPath);
            XmlNodeType type;
            while (reader.Read())
            {
                type = reader.NodeType;
                if (!type.Equals(XmlNodeType.Element))
                {
                    result.AddRange(reader.Value.ToString().Split(new string[] { "\t\r\n", " " }, StringSplitOptions.RemoveEmptyEntries));
                }
            }

            listresult.Add(filenameWithPath, result);
            return listresult;
        }

        /// <summary>
        /// Reads HTML file types and extracts the words in each HTML file
        /// Requires: The file path is in .html format only
        /// </summary>
        /// <param name="filenameWithPath">path of HTML document including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when the file to read is not of supported
        /// html format.</exception>
        /// <returns>
        /// A Dictionary where the Key contains the filename and the Value contains the entire wordlist 
        /// </returns>
        private static Dictionary<string, List<string>> readHTMLFiles(string filenameWithPath)
        {
            List<string> result = new List<string>();
            Dictionary<string, List<string>> listresult = new Dictionary<string, List<string>>();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(filenameWithPath);
            var myHTMLNodes = doc.DocumentNode.SelectNodes("//text()");
            foreach (HtmlNode node in myHTMLNodes)
            {
                result.AddRange(node.InnerText.Split(new string[] { "\t\r\n", " " }, StringSplitOptions.RemoveEmptyEntries));
            }
            listresult.Add(filenameWithPath, result);
            return listresult;
        }

        /// <summary>
        /// Action to read file of different extensions by detecting extensions
        /// Can read only .pdf, .doc, .docx, .ppt, .pptx, .xls, .xlsx, .txt .html and .xml 
        /// 
        /// </summary>
        /// <param name="filenameWithPath">document path including filename</param>
        /// <exception cref="PlatformNotSupportedException">Thrown when there is an attempt to read
        /// a file with an unsupported file extension</exception>
        /// <returns>the list of tokens for each file</returns>
        public static Dictionary<string, List<string>> readFiles(string filenameWithPath)
        {
           string extension = System.IO.Path.GetExtension(filenameWithPath).ToLower();
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();

            switch (extension.Replace(".",""))
                {
                    case ("xml"):
                        result = readXMLFiles(filenameWithPath);
                        break;

                    case ("xls"):
                    case ("xlsx"):

                        result = readExcelFiles(filenameWithPath,extension);
                        break;
                    case ("doc"):
                    case ("docx"):
                        result = readDocFiles(filenameWithPath);
                        break;

                    case ("txt"):
                        result = readTxtFiles(filenameWithPath);
                        break;

                    case ("pdf"):
                        result = readPdfFile(filenameWithPath);
                        break;

                    case ("html"):
                        result = readHTMLFiles(filenameWithPath);
                        break;

                    case ("ppt"):
                    case ("pptx"):
                        result = readPPTFiles(filenameWithPath);
                        break;

                    default:
                        break;
                }

                //This is passed to the word dictionary to print out the words in the file
                WordDictionary.SingleWordDictionary(result);

                return result;
            }
            
        

        /// <summary>
        /// Attemps to read all the files in a given directory and return a list of tokens for each file
        /// </summary>
        /// <param name="directoryname">directory to read from</param>
        /// <returns>a list of tokens for each files</returns>
        public static List<Dictionary<string, List<string>>> readfromDirectory(string directoryname)
        {
            List<Dictionary<string, List<string>>> result = new List<Dictionary<string, List<string>>>();
            foreach (string dir in Directory.GetDirectories(directoryname))
            {

                foreach (var document in Directory.GetFiles(directoryname))
                {
                    //This is the name of the document to search now
                   // MessageBox.Show(document.ToString());
                    result.Add(readFiles(document));
                }
                readfromDirectory(dir);
            }

            return result;

        }

    }
}
