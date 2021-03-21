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
        SharePointManager spManager = null;
        string result = string.Empty;

        //StringBuilder testlibName = new StringBuilder("testLib");
        public Form1(SharePointManager spManager)
        {
            this.spManager = spManager;
            InitializeComponent();
            //spManager.ConnectToSharePoint("nehad.goher@mylinkdev.onmicrosoft.com","Neh@d123",out result);
            this.richTextBox1.AppendText(result);


        }
        string specifyList = "As_Is_Documents_Reference";
        string sheetName = "Sheet1";
        int rowStart = 2;
        // browse button
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK && !String.IsNullOrWhiteSpace(openFileDialog2.FileName))
            {
                this.richTextBox1.AppendText("Opening file .....\n");
                excel = new ExcelFileManager(openFileDialog2.FileName);
                string res = excel.OpenFileExcel();
                this.richTextBox1.AppendText(res);
                //PreProcessing(2,"Sheet1",)
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if(excel != null)
            {
                excel.ReleaseFileResources();
            }


        }
        /// Delete button
        private void btn_loadList_Click(object sender, EventArgs e)
        {
            this.richTextBox1.AppendText("Start Processing .... \n");
            if (excel != null)
            {
                Excel.Worksheet xlWorkSheet = excel.GetExcelSheetByName("Sheet1");
                this.richTextBox1.AppendText("retriving the data of the sheet .... \n");
                DeleteItems(xlWorkSheet);
                this.richTextBox1.AppendText("Finished ----------------------------\n");
            }
            else
            {
                MessageBox.Show("Must select valid Excel file");
            }
         
        }
        private void DeleteItems(Excel.Worksheet xlWorkSheet)
        {
            int totalDeleted = 0,  total = excel.GetLastRowInSheet(xlWorkSheet);
            /// if any result happens during retriving the data
            this.richTextBox1.AppendText(result);
            SP.List lib = spManager.GetListByName(specifyList,out result);
            this.richTextBox1.AppendText(result);
            if (lib!= null)
            {
                for (int iRow = rowStart; iRow <= total; iRow++)
                {
                    if(xlWorkSheet.Cells[iRow, 1].value != null)
                    {
                        string item = (string)xlWorkSheet.Cells[iRow, 1].value;

                        var success = spManager.DeleteItemFromFolder(lib, item.Trim(), out result);
                        if (success) totalDeleted++;
                        this.richTextBox1.AppendText(result);
                    }
                }
                this.richTextBox1.AppendText($"Deleted : {totalDeleted} / { total-1}\n");
            }
        }
    }
}
