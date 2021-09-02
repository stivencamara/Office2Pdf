using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using System;
/*
https://github.com/skmp/node-msoffice-pdf/blob/master/office.cs
*/

namespace Office2Pdf.Converters
{
    internal class PowerpointConverter : ConverterBase<Application>
    {
        public PowerpointConverter() :
            base(new Application())
        {
        }
        
        public override void OnConvert(string sourcePath, string targetPath, bool isPdfA)
        {
            object unknownType = Type.Missing;

            Application.WindowState = PpWindowState.ppWindowMinimized;
            Application.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            Application.Visible = MsoTriState.msoTrue;//Not is permitted hidden
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var presentations = Application.Presentations;
            var presentation = presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

            try
            {
                var printRange = presentation.PrintOptions.Ranges.Add(1, presentation.Slides.Count);

                presentation.ExportAsFixedFormat(targetPath, PpFixedFormatType.ppFixedFormatTypePDF, 
                    PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, printRange, PpPrintRangeType.ppPrintAll,
                    unknownType, true, true, true, true, isPdfA, unknownType);
            }
            catch(Exception e) 
            {
                try
                {
                    presentation.SaveAs(targetPath, PpFixedFormatType.ppFixedFormatTypePDF, MsoTriState.msoCTrue);
                }
                catch 
                { }
            }
            finally
            {
                if (Application != null)
                {
                    if (Application.Presentations.Count > 0)
                        presentation.Close();
                }
                
                Application.Quit();
                Application.Dispose();
            }

        }
    }
}
