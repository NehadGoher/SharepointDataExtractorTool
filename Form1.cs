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

        Action<int, Excel.Worksheet> ProcessingMethod;

        StringBuilder testlibName = new StringBuilder("testLib");
        public Form1(SharePoitnOnlineManager spManager)
        {
            this.spManager = spManager;
            InitializeComponent();
            //this.bindingNavigator1.BindingSource = new 
            //spManager.ConnectToSharePoint("nehad.goher@mylinkdev.onmicrosoft.com","Neh@d123",out result);
            this.richTextBox1.AppendText(result);
        }

        // browse
        private void button1_Click(object sender, EventArgs e)
        {
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
        }
        private void CreateContentTypes(int rowStart, Excel.Worksheet xlWorkSheet)
        {
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
                    string contentTypeame = (string)xlWorkSheet.Cells[iRow, 1].value;
                    string parentName = "Item";
                    if (iRow > 3)
                    {
                        parentName = (string)xlWorkSheet.Cells[iRow, 2].value;
                    }
                    result = spManager.CreateContentType(contentTypeame, parentName);
                    this.richTextBox1.AppendText(result);
                }
            }
        }

        private void CreateColumnsInContentType(int rowStart, Excel.Worksheet xlWorkSheet)
        {

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
                    result = spManager.CreateSiteColumn(name.Trim(), contentType.Trim(), type.Trim(), isRequired, "DFTC Site Column");
                    this.richTextBox1.AppendText(result);
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

        private void test_btn_Click(object sender, EventArgs e)
        {
            if(excel != null)
            {
                this.richTextBox1.AppendText("Loading data to show into the lists\n");
                this.listContent.DataSource = LoadDataFromSheet(3,1, "Library Mapping");
                this.listSite.DataSource = LoadDataFromSheet(2,2, "DFTC Content Types");
                this.listLib.DataSource = LoadDataFromSheet(12,2, "Library Mapping");
                this.richTextBox1.AppendText("Finnished ----------\n");
            }
            else
            {
                MessageBox.Show("Must select excel file");
            }
          
        }

        private void Form1_Closing(object sender, CancelEventArgs e)
        {
            excel.ReleaseFileResources();
        }

        private void bt_createContentType_Click(object sender, EventArgs e)
        {
            //if ()
            //{
                PreProcessing(3, "Library Mapping", CreateContentTypes);
            //}
            
        }

        private void btn_createSiteColumn_Click(object sender, EventArgs e)
        {
            PreProcessing(2, "DFTC Content Types", CreateColumnsInContentType);
        }

        private void btn_createLibrary_Click(object sender, EventArgs e)
        {
            PreProcessing(12, "Library Mapping", CreateLibrary);
        }

        private void btn_deleteContentTypes_Click(object sender, EventArgs e)
        {
            PreProcessing(3, "Library Mapping", RemoveContentTypes);
        }

        private void btn_deleteSiteColumn_Click(object sender, EventArgs e)
        {
            PreProcessing(2, "DFTC Content Types", RemoveSiteColmun);
        }

        private void btn_deleteLibrary_Click(object sender, EventArgs e)
        {
            PreProcessing(13, "Library Mapping", RemoveLibrary);
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

       

        //// load data
        //private void button2_Click(object sender, EventArgs e)
        //{
        //    // must select sheet name
        //    string SheetName = this.comboBox1.GetItemText(this.comboBox1.SelectedItem?.ToString() ?? "");
        //    if (SheetName.Length > 0)
        //    {
        //        Excel.Worksheet  xlWorkSheet = excel.GetExcelSheetByName(SheetName);
        //        this.richTextBox1.AppendText("retriving the data of the sheet .... \n");
        //        int rowStart;
        //        /// create content type
        //        if (SheetName.Contains("Library"))
        //        {
        //            rowStart = 3;
        //            CreateContentTypes(rowStart, xlWorkSheet);
        //        }
        //        else
        //        {
        //            /// create column in content type
        //            rowStart = 2;
        //            CreateColumnsInContentType(rowStart, xlWorkSheet);
        //        }
        //        this.richTextBox1.AppendText("Finished ----------------------------\n");
        //    }
        //    else
        //    {
        //        MessageBox.Show("Must Select Sheet Name from the excel file");
        //    }
        //}

        //private void deleteDataBtn_Click(object sender, EventArgs e)
        //{
        //    string SheetName = this.comboBox1.GetItemText(this.comboBox1.SelectedItem?.ToString() ?? "");
        //    if (SheetName.Length > 0)
        //    {
        //        this.richTextBox1.AppendText("Reteriving required data.... \n");
        //        Excel.Worksheet xlWorkSheet = excel.GetExcelSheetByName(SheetName);
        //        int rowStart;
        //        /// create content type
        //        if (SheetName.Contains("Library"))
        //        {
        //            rowStart = 3;
        //            RemoveContentTypes(rowStart, xlWorkSheet);
        //        }
        //        else
        //        {
        //            /// create column in content type
        //            rowStart = 2;
        //            RemoveSiteColmun(rowStart, xlWorkSheet);
        //        }
        //        this.richTextBox1.AppendText("Finished ----------------------------\n");
        //    }
        //    else
        //    {
        //        MessageBox.Show("Must Select Sheet Name from the excel file");
        //    }
        //}
    }
}
