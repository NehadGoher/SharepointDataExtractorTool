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
        string error = string.Empty;

        Action<int, Excel.Worksheet> ProcessingMethod;

        StringBuilder testlibName = new StringBuilder("testLib");
        public Form1(SharePoitnOnlineManager spManager)
        {
            this.spManager = spManager;
            InitializeComponent();
            //spManager.ConnectToSharePoint("nehad.goher@mylinkdev.onmicrosoft.com","Neh@d123",out error);
            this.richTextBox1.AppendText(error);
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
            List<string> ContentTypes = spManager.GetContentTypesName(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);
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
                    error = spManager.CreateContentType(contentTypeame, parentName);
                    this.richTextBox1.AppendText(error);
                }
            }
        }

        private void CreateColumnsInContentType(int rowStart, Excel.Worksheet xlWorkSheet)
        {

            List<string> Fields = spManager.GetSiteColumnsName(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);

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
                    error = spManager.CreateSiteColumn(name.Trim(), contentType.Trim(), type.Trim(), isRequired, "DFTC Site Column");
                    this.richTextBox1.AppendText(error);
                }
            }

        }

        /// still under testing
        private void CreateLibrary(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            List<string> lists = spManager.GetLists(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);
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
                    error = spManager.AddContentTypeToListByNames(contentType.Trim(),name.Trim());
                    this.richTextBox1.AppendText(error);
                }
            }

        }

        private void RemoveSiteColmun(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            this.richTextBox1.AppendText("Deleting .....\n");
            List<string> ContentTypes = spManager.GetContentTypesName(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);
            List<string> Fields = spManager.GetSiteColumnsName(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);
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
                    error = spManager.DeleteSiteColumnFromContentType(Regex.Replace(contentType, @"\s+", "").Trim(), Regex.Replace(name, @"\s+", "").Trim(), type);
                    this.richTextBox1.AppendText(error);
                }
            }

        }

        private void RemoveContentTypes(int rowStart, Excel.Worksheet xlWorkSheet)
        {
            this.richTextBox1.AppendText("Deleting .....\n");
            List<string> ContentTypes = spManager.GetContentTypesName(out error);
            /// if any error happens during retriving the data
            this.richTextBox1.AppendText(error);
            for (int iRow = rowStart; iRow <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    string contentTypeName = (string)xlWorkSheet.Cells[iRow, 1].value;
                    
                    error = spManager.DeleteContentTypesWithName(Regex.Replace(contentTypeName, @"\s+", "").Trim());
                    this.richTextBox1.AppendText(error);
                }
            }
        }

        private void test_btn_Click(object sender, EventArgs e)
        {
            testlibName.Append(Guid.NewGuid());
            spManager.AddContentTypeToListByNames(testlibName.ToString(),"TestList");
        }

        private void Form1_Closing(object sender, CancelEventArgs e)
        {
            excel.ReleaseFileResources();
        }

        private void bt_createContentType_Click(object sender, EventArgs e)
        {
            ProcessingMethod = CreateContentTypes;
            PreProcessing(3, "Library Mapping", ProcessingMethod);
        }

        private void btn_createSiteColumn_Click(object sender, EventArgs e)
        {
            ProcessingMethod = CreateColumnsInContentType;
            PreProcessing(2, "DFTC Content Types", ProcessingMethod);
        }

        private void btn_createLibrary_Click(object sender, EventArgs e)
        {

            ProcessingMethod = CreateContentTypes;
            PreProcessing(12, "DFTC Content Types", ProcessingMethod);
        }

        private void btn_deleteContentTypes_Click(object sender, EventArgs e)
        {
            ProcessingMethod = RemoveContentTypes;
            PreProcessing(3, "Library Mapping", ProcessingMethod);
        }

        private void btn_deleteSiteColumn_Click(object sender, EventArgs e)
        {
            ProcessingMethod = RemoveSiteColmun;
            PreProcessing(2, "DFTC Content Types", ProcessingMethod);
        }

        private void btn_deleteLibrary_Click(object sender, EventArgs e)
        {
            ProcessingMethod = RemoveSiteColmun;
           // PreProcessing(13, "Library Mapping", ProcessingMethod);
           
        }

        private void PreProcessing (int row,string sheetName, Action<int, Excel.Worksheet> creation)
        {
            this.richTextBox1.AppendText("Start processing ......\n");
            if (excel != null)
            {
                Excel.Worksheet xlWorkSheet = excel.GetExcelSheetByName(sheetName);
                this.richTextBox1.AppendText("retriving the data of the sheet .... \n");
                creation(row, xlWorkSheet);
                this.richTextBox1.AppendText("Finished ----------------------------\n");
            }
            else
            {
                MessageBox.Show("Must Select file Excel");
            }
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
