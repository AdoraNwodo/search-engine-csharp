using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchEngine
{
    /// <summary>
    /// Search Engine Home page
    /// </summary>
    public partial class frmHomePage : Form
    {
        
        // The directory path
        string directoryPath = @"C:\Users\USER\Dropbox\Search Engine_Updated (1)\Search Engine_Updated\SearchEngine\files";
        

        /// <summary>
        /// Search Engine Home page
        /// </summary>
        public frmHomePage()
        {
            InitializeComponent();
            
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            try
            {
                lvSearchResult.Clear();
                DirectoryOperations dir = new DirectoryOperations(directoryPath);
                DirectoryOperations.readfromDirectory(directoryPath);
                 // Search all files in the directory for the entered text


                List<string> lstRelevantFilePaths = dir.Find(txtSearch.Text).TopDocuments;
                double time = dir.Find(txtSearch.Text).ResponseTime;

                  // Display a decsription for the search result
                if (lstRelevantFilePaths.Count != 0)
                {
                    // Some files were returned
                    lblSearchResult.Text = lstRelevantFilePaths.Count + " File(s) Found " + time + "milli-seconds";
                }
                else
                {
                    // No file was found
                    lblSearchResult.Text = "No File Found, " + time + " milli-seconds";
                }

                   // Display the search result
                foreach (string filePath in lstRelevantFilePaths)
                {
                    lvSearchResult.Items.Add(filePath);
                    
                    lvSearchResult.Items.Add("");
                    lvSearchResult.FullRowSelect = true;
                    this.lvSearchResult.Click += new EventHandler(this.lvSearchResult_Click);
                }
            }
            catch (Exception) { }

        }

        private void lvSearchResult_Click(object sender, EventArgs e)
        {
            //action taken when a search result in the list view item is clicked on.
            if (lvSearchResult.SelectedItems.Count > 0)
                Process.Start("explorer.exe", " /open, " + lvSearchResult.SelectedItems[0].Text);
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DatabaseEngine d = new DatabaseEngine();

                // Create the list to use as the custom source. 
                var source = new AutoCompleteStringCollection();
                List<string> terms = d.GetSearchedTerms();
                foreach (string term in terms)
                {
                    source.Add(term);
                }
                txtSearch.AutoCompleteCustomSource = source;
                txtSearch.AutoCompleteMode = AutoCompleteMode.Suggest;
                txtSearch.AutoCompleteSource = AutoCompleteSource.CustomSource;

            }
            catch (Exception) { }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void frmHomePage_Load(object sender, EventArgs e)
        {

        }

        
    }
}
