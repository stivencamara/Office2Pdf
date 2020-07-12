
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System;

namespace Office2Pdf.Converters
{
    internal class WordConverter : ConverterBase<Application>
    {
        private static object unknownType = Type.Missing;
        public WordConverter() :
            base(new Application())
        {

        }
        public override void OnConvert(string sourcePath, string targetPath, bool isPdfA)
        {
            Application.Visible = false;

            try
            {
                Documents documents = Application.Documents;

                documents.Open(sourcePath, unknownType, true, unknownType, unknownType, unknownType,
                    unknownType, unknownType, unknownType, unknownType, unknownType, unknownType,
                    unknownType, unknownType, unknownType, unknownType);


                Application.Application.Visible = false;
                Application.WindowState = WdWindowState.wdWindowStateMinimize;

                Document activeDocument = Application.ActiveDocument;

                if (isPdfA)
                {
                    try
                    {
                        activeDocument.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF, false,
                            WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent,
                            true, true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, true);
                    }
                    catch
                    {
                        SaveAs(targetPath, activeDocument);
                    }
                }
                else
                {
                    SaveAs(targetPath, activeDocument);
                }
            }
            finally
            {
                if (Application != null)
                {
                    if (Application.Documents.Count > 0)
                        Application.Documents.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdWordDocument);
                }
                Application.Quit(unknownType, unknownType, unknownType);
            }
        }

        private void SaveAs(string targetPath, Document activeDocument)
        {
            object fileFormat = WdSaveFormat.wdFormatPDF;
            object fileName = targetPath;
            activeDocument.SaveAs(fileName, fileFormat, unknownType, unknownType, unknownType,
                unknownType, unknownType, unknownType, unknownType, unknownType, unknownType,
                unknownType, unknownType, unknownType, unknownType, unknownType);
        }
    }
}
