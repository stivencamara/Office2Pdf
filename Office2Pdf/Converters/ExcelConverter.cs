using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.OfficeApi.Enums;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Office2Pdf.Converters
{
    internal class ExcelConverter : ConverterBase<Application>
    {
        public ExcelConverter() :
            base(new Application())
        {
        }

        public override void OnConvert(string sourcePath, string targetPath, bool isPdfA)
        {
            Application.Visible = false;
            Application.ScreenUpdating = false;
            Application.DisplayAlerts = false;
            Application.Application.Visible = false;
            Application.WindowState = XlWindowState.xlMinimized;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            Application.AskToUpdateLinks = false;

            Workbooks workbooks = null;
            Workbook workbook = null;

            try
            {
                object unknownType = Type.Missing;
                workbooks = Application.Workbooks;
                workbook = workbooks.Open(sourcePath);

                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetPath);
                workbook.Close(false, unknownType, unknownType);

                workbooks.Close();
                workbook.DisposeChildInstances();
            }
            finally
            {
                if (Application != null)
                {
                    if (Application.Workbooks.Count > 0)
                        Application.Workbooks.Close();
                }
                Application.Quit();
                Application.Dispose();
                //ForceQuit(workbooks, workbook);
            }
        }

        //[DllImport("user32.dll")]
        //private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        //private void ForceQuit(Workbooks workbooks, Workbook workbook)
        //{
        //    int hWnd = Application.Application.Hwnd;
        //    uint processID;

        //    GetWindowThreadProcessId((IntPtr)hWnd, out processID);
        //    Process.GetProcessById((int)processID).Kill();

        //    workbooks = null;
        //    workbook = null;
            
        //}
    }
}
