using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SP = Microsoft.SharePoint.Client;
using System.Security;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ContentTypeExtractor
{
    public partial class Form1 : Form
    {
        ExcelFileManager excel = null;
        SharePoitnOnlineManager spManager = null;
        string result = string.Empty;

        const string ContentTypeSheetName = "Library Mapping";
        const string LibrarySheetName = "Library Mapping";
        const string SiteColumnSheetName = "DFTC Content Types";

       const int ContentStartRow = 3;
       const int SiteStartRow = 2;
       const int LibraryStartRow = 12;

        public Form1(SharePoitnOnlineManager spManager)
        {
            this.spManager = spManager;
            InitializeComponent();
            this.listLib.Enabled = false;
            this.listSite.Enabled = false;
            this.richTextBox1.AppendText(result);
        }

        // browse
        private void button1_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            if (openFileDialog2.ShowDialog() == DialogResult.OK && !String.IsNullOrWhiteSpace(openFileDialog2.FileName))
            {
                this.richTextBox1.AppendText("Opening file .....\n");
                excel = new ExcelFileManager(openFileDialog2.FileName);
                string res = excel.OpenFileExcel();
                this.richTextBox1.AppendText(res);
            }
            else
            {
                this.richTextBox1.AppendText("File Dialog can't be opened \n");
            }
            ToggleAllButtons(true);
        }
        private void CreateContentTypes(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            List<string> ContentTypes = spManager.GetContentTypesName(out result);
            string selectedContentType = this.listContent.SelectedItem?.ToString()??"".Trim();
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentTypeName = (string)xlWorkSheet.Cells[iRow, 1].value;
                    string parentName = "Item";
                    if (iRow > 3)
                    {
                        parentName = (string)xlWorkSheet.Cells[iRow, 2].value;
                    }
                    if (  string.IsNullOrWhiteSpace(selectedContentType)|| contentTypeName.Trim() == selectedContentType)
                    {
                        result = spManager.CreateContentType(contentTypeName, parentName);
                        this.richTextBox1.AppendText(result);
                    }
                    
                }
            }
            if (!string.IsNullOrWhiteSpace(selectedContentType))
            {
                CreateColumnsInContentType(2, excel.GetExcelSheetByName("DFTC Content Types"));
            }
        }

        private void CreateColumnsInContentType(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            string selectedContentType = this.listContent.SelectedItem?.ToString() ?? "";
            List<string> Fields = spManager.GetSiteColumnsName(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);

            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentType = (string)xlWorkSheet.Cells[iRow, 1].value;
                    string name = (string)xlWorkSheet.Cells[iRow, 2].value;
                    string type = (string)xlWorkSheet.Cells[iRow, 3].value;
                    bool isRequired = ((string)xlWorkSheet.Cells[iRow, 4].value == "Mandatory") ? true : false;
                    if (string.IsNullOrWhiteSpace(selectedContentType) || contentType.Trim() == selectedContentType)
                    {
                        result = spManager.CreateSiteColumn(name.Trim(), contentType.Trim(), type.Trim(), isRequired, "DFTC Site Column");
                        this.richTextBox1.AppendText(result);
                    }
                        
                }
            } 

        }

        private void CreateLibrary(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            List<string> lists = spManager.GetLists(out result);
            this.richTextBox1.AppendText(result);
            spManager.GetContentTypesName(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string name = (string)xlWorkSheet.Cells[iRow, 1].value;
                    string contentType = (string)xlWorkSheet.Cells[iRow, 2].value;
                  
                    result = spManager.CreateLibrary(name.Trim());
                    this.richTextBox1.AppendText(result);
                    result = spManager.AddContentTypeToListByNames(contentType.Trim(),name.Trim());
                      this.richTextBox1.AppendText(result);
                }
            }

        }

        private void RemoveSiteColmun(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            this.richTextBox1.AppendText("Deleting .....\n");
            List<string> ContentTypes = spManager.GetContentTypesName(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            List<string> Fields = spManager.GetSiteColumnsName(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentType = (string)xlWorkSheet.Cells[iRow, 1].value;
                    string name = (string)xlWorkSheet.Cells[iRow, 2].value;
                    string type = (string)xlWorkSheet.Cells[iRow, 3].value;
                    bool isRequired = ((string)xlWorkSheet.Cells[iRow, 4].value == "Mandatory") ? true : false;
                    result = spManager.DeleteSiteColumnFromContentType(Regex.Replace(contentType, @"\s+", "").Trim(), Regex.Replace(name, @"\s+", "").Trim(), type);
                    this.richTextBox1.AppendText(result);
                }
            }

        }

        private void RemoveContentTypes(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            this.richTextBox1.AppendText("Deleting .....\n");
            List<string> ContentTypes = spManager.GetContentTypesName(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentTypeName = (string)xlWorkSheet.Cells[iRow, 1].value;
                    
                    result = spManager.DeleteContentTypesWithName(Regex.Replace(contentTypeName, @"\s+", "").Trim());
                    this.richTextBox1.AppendText(result);
                }
            }
        }

        private void RemoveLibrary(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            this.richTextBox1.AppendText("Deleting .....\n");
            List<string> Lst = spManager.GetLists(out result);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string lib = (string)xlWorkSheet.Cells[iRow, 1].value;

                    result = spManager.DeleteList(Regex.Replace(lib, @"\s+", "").Trim());
                    this.richTextBox1.AppendText(result);
                }
            }
        }
        /// load data
        private void test_btn_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            if (excel != null)
            {
                this.richTextBox1.AppendText("Loading data to show into the lists\n");
                this.listContent.DataSource = LoadDataFromSheet(ContentStartRow,1, ContentTypeSheetName);
                this.listSite.DataSource = LoadDataFromSheet(SiteStartRow,2, SiteColumnSheetName);
                this.listLib.DataSource = LoadDataFromSheet(LibraryStartRow,1, LibrarySheetName);
                ClearAllSelected();
                this.richTextBox1.AppendText("Finnished ----------\n");
            }
            else
            {
                MessageBox.Show("Must select excel file");
            }
            ToggleAllButtons(true);
        }

        private void bt_createContentType_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(ContentStartRow, ContentTypeSheetName, CreateContentTypes);
            ToggleAllButtons(true);
        }

        private void btn_createSiteColumn_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(SiteStartRow, SiteColumnSheetName, CreateColumnsInContentType);
            ToggleAllButtons(true);
        }

        private void btn_createLibrary_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(LibraryStartRow, LibrarySheetName, CreateLibrary);
            ToggleAllButtons(true);
        }

        private void btn_deleteContentTypes_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(ContentStartRow, ContentTypeSheetName, RemoveContentTypes);
            ToggleAllButtons(true);
        }

        private void btn_deleteSiteColumn_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(SiteStartRow, SiteColumnSheetName, RemoveSiteColmun);
            ToggleAllButtons(true);
        }

        private void btn_deleteLibrary_Click(object sender, EventArgs e)
        {
            ToggleAllButtons();
            PreProcessing(LibraryStartRow, LibrarySheetName, RemoveLibrary);
            ToggleAllButtons(true);
        }

        private void PreProcessing (int row,string sheetName, Action<int, Excel.Worksheet> processing)
        {
            this.richTextBox1.AppendText("Start processing ......\n");
            if (excel != null)
            {
                Excel.Worksheet xlWorkSheet = excel.GetExcelSheetByName(sheetName);
                this.richTextBox1.AppendText("retriving the data of the sheet .... \n");
                processing(row, xlWorkSheet);
                this.richTextBox1.AppendText("Finished ----------------------------\n");
            }
            else
            {
                MessageBox.Show("Must Select file Excel");
            }
        }

        private void btn_clear_Click(object sender, EventArgs e)
        {
            ClearAllSelected();
        }

        private void ClearAllSelected()
        {
            this.listContent.SelectedIndices.Clear();
            this.listSite.SelectedIndices.Clear();
            this.listLib.SelectedIndices.Clear();
        }
        private List<string> LoadDataFromSheet(int startRow,int colVal ,string sheetName)
        {
            List<string> data = new List<string>();
            var xlWorkSheet = excel.GetExcelSheetByName(sheetName);
            for (int iRow = startRow; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, colVal].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentTypeName = (string)xlWorkSheet.Cells[iRow, colVal].value;
                    data.Add(contentTypeName);
                }
            }
            return data;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if(excel != null)
            {
                excel.ReleaseFileResources();
            }
                spManager.Dispose();
        }

        private void ToggleAllButtons(bool flag = false)
        {
            this.btn_clear.Enabled = this.btn_createLibrary.Enabled = this.btn_createSiteColumn.Enabled
                = this.btn_deleteContentTypes.Enabled = this.btn_deleteLibrary.Enabled = this.btn_deleteSiteColumn.Enabled
                = this.test_btn.Enabled = this.bt_createContentType.Enabled =this.buttonBrowse.Enabled = flag;
        }
    }
}
