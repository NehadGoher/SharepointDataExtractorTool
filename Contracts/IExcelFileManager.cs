using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContentTypeExtractor.Contracts
{
    interface IExcelFileManager
    {
        string OpenFileExcel();
        List<string> GetSheetsName();
        Excel.Worksheet GetExcelSheetByName(string name);
        bool ReleaseFileResources();
        int GetLastRowInSheet(Excel.Worksheet sheet);
    }
}
