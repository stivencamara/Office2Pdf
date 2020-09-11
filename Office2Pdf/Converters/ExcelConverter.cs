using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.OfficeApi.Enums;
using System;

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

            try
            {
                object unknownType = Type.Missing;
                var workbooks = Application.Workbooks;
                var workbook = workbooks.Open(sourcePath);

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
            }
        }
    }
}
