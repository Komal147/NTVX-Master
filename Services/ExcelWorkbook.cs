using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace NTVXApi.Services
{
    public sealed class ExcelWorkbook: IDisposable
    {
        Application excel = null;
        Workbooks workbooks = null;
        Workbook workbook = null;
        Worksheet worksheet = null;

        public Application ExcelApplication => excel;

        public void Open(string excelFile)
        {
            excel = new Application();
            excel.Visible = false;
            workbooks = excel.Workbooks;

            workbook = workbooks.Open(excelFile);
        }

        public Worksheet OpenWorksheet(string sheetName)
        {
            worksheet = (Worksheet)workbook.Sheets[sheetName];
            return worksheet;
        }

        public string RunMacro(string macroName, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            macroName = "'" + workbook.Name + "'!" + macroName;
            string ret = excel.Run(macroName, arg1, arg2, arg3, arg4, arg5, arg6);
            if (ret.Substring(0, 6) == "Error:")
                throw new ApplicationException(ret);
            workbook.Save();
            
            return ret;
        }

        public void Dispose()
        {
            if (excel != null)
            {
                if (excel.ActiveWindow != null)
                    excel.ActiveWindow.Close(true);
                excel.Quit();
            }

            FinalReleaseComObject(worksheet);
            FinalReleaseComObject(workbook);
            FinalReleaseComObject(workbooks);
            FinalReleaseComObject(excel);

            GC.SuppressFinalize(this);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static void FinalReleaseComObject(object comObj)
        {
            if (comObj != null)
            {
                Marshal.FinalReleaseComObject(comObj);
                comObj = null;
            }
        }
    }
}
