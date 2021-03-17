using ContentTypeExtractor.Contracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContentTypeExtractor
{
    public class ExcelFileManager : IExcelFileManager
    {

        string filePath = String.Empty;
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        public ExcelFileManager(string filePath)
        {
            this.filePath = filePath;
        }

        private bool CheckFileExtention()
        {
            string fileExt = string.Empty;  //get the path of the file  
            fileExt = Path.GetExtension(filePath); //get the file extension  

            if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// open file excel with path in ctor
        /// </summary>
        /// <returns> return what happens</returns>
        public string OpenFileExcel()
        {
            try
            {
                if (CheckFileExtention())
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(filePath);
                    return $"{filePath} opend successfully \n";
                }
                else
                {
                    return "file extention must be .xlsx or .xls \n";
                }
            }
            catch (Exception ex)
            {
                xlWorkBook.Close();
                xlApp.Quit();
                return $"{ex.Message.ToString()} in {MethodBase.GetCurrentMethod()}\n";
            }
        }

        /// <summary>
        /// get sheets names
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetsName()
        {
            List<string> sheets = new List<string>();
            try {
                foreach (Excel.Worksheet ws in xlWorkBook.Worksheets)
                {
                    sheets.Add(ws.Name);
                }
                return sheets;
            } catch (Exception ex)
            {
                return sheets;
            }
        }

        /// <summary>
        ///  get sheet data by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns> the whole data of sheet by name</returns>
        public Excel.Worksheet GetExcelSheetByName(string name) => xlWorkBook.Worksheets[name];

        /// <summary>
        /// release and close the file 
        /// </summary>
        /// <returns>success or fail</returns>
        public bool ReleaseFileResources()
        {
            try
            {
                xlWorkBook.Close();
                xlApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// get real number of the row in the sheet by sheet object
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public int GetLastRowInSheet(Excel.Worksheet sheet)  =>
                               sheet.Cells.Find("*", Missing.Value,Missing.Value, Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, Missing.Value, Missing.Value).Row;
    
    }
}
