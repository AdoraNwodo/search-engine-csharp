using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SearchEngine;
using System.Collections.Generic;

namespace SearchTest
{
    [TestClass]
    public class UnitTest1
    {
        private static string dirForRead = @"C:\Users\USER\Desktop\SearchEngine\files\";
           
        [TestMethod]
        public void TestForTXTFiles()
        {
            string path = dirForRead+"TextFile.txt";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "This";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForXLSFiles()
        {

            string path = dirForRead + "Book1.xls";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "adaora";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForDOCXFiles()
        {

            string path = dirForRead + "CSC 307 project.docx";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "CSC";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForDOCFiles()
        {

            string path = dirForRead + "project3.doc";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "Extend";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForXLSXFiles()
        {

            string path = dirForRead + "ExampleData.xlsx";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "dangers";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForPDFFiles()
        {

            string path = dirForRead + "CSC 307 project.pdf";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "csc";
            Assert.AreEqual(expected, actual);
        }

        
        [TestMethod]
        public void TestForHTMLFiles()
        {

            string path = dirForRead + "WelcomePage.html";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "Welcome";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestForXMLFiles()
        {

            string path = dirForRead + "MyXml.xml";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "mmm";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadTXTFiles2()
        {

            string path = dirForRead + "TextFile.txt";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "file";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadXLSFiles2()
        {

            string path = dirForRead + "Book1.xls";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "oluwatobi";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadDOCXFiles2()
        {

            string path = dirForRead + "CSC 307 project.docx";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "this";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadDOCFiles2()
        {

            string path = dirForRead + "day.doc";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "lord";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadXLSXFiles2()
        {

            string path = dirForRead + "ExampleData.xlsx";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "2";
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReadPDFFiles2()
        {

            string path = dirForRead + "day.pdf";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][output[path].Count - 1];
            string expected = "lord";
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void ReadPPTFiles()
        {

            string path =  dirForRead+"lec18.ppt";
            Dictionary<string, List<string>> output = FileHandler.readFiles(path);
            string actual = output[path][0];
            string expected = "mathematical";
            Assert.AreEqual(expected, actual);
        }

        
        [TestMethod]
        [ExpectedException(typeof(PlatformNotSupportedException))]
        public void TestWithWrongFormatCSV()
        {
            string path = @"C:\Users\USER\Desktop\CSC 322- GROUP 4\Search Engine_Updated (1)\Search Engine_Updated\SearchEngine\wrong format test file\Book1.csv";
            Dictionary<string, List<string>> output = FileHandler.readExcelFiles(path, ".csv");

        }

        [TestMethod]
        [ExpectedException(typeof(PlatformNotSupportedException))]
        public void TestWithWrongFormatRTF()
        {
            string path = @"C:\Users\USER\Desktop\CSC 322- GROUP 4\Search Engine_Updated (1)\Search Engine_Updated\SearchEngine\wrong format test file\ReadMe.rtf";
            Dictionary<string, List<string>> output = FileHandler.readDocFiles(path);

        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestFindWithEmptyQuery()
        {
            DirectoryOperations dirOP = new DirectoryOperations(dirForRead);
            string query = "";
            List<string> output = dirOP.Find(query).TopDocuments;

        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestFindWithWhitespaceQuery()
        {
            DirectoryOperations dirOP = new DirectoryOperations(dirForRead);
            string query = "    ";
            List<string> output = dirOP.Find(query).TopDocuments;

        }
        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestNullInputInSortedDictionary()
        {
            SortedDictionary<string, Keywords> output = new SortedDictionary<string, Keywords>();
            Dictionary<string, List<string>> listresult = null;
            output = WordDictionary.SingleWordDictionary(listresult);
          
        }
                
    }
}
