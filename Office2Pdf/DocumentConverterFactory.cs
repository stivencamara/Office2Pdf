using Office2Pdf.Converters;
using System;
using System.Collections.Generic;
using System.Text;

namespace Office2Pdf
{
    public class DocumentConverterFactory : IDocumentConverterFactory
    {
        public IConverter GetConverter(ContentType contentType)
        {
            IConverter converter = null;

            switch (contentType)
            {
                case ContentType.DOC:
                case ContentType.DOCX:
                    converter = new WordConverter();
                    break;
                case ContentType.XLS:
                case ContentType.XLSX:
                    converter = new ExcelConverter();
                    break;
                case ContentType.PPT:
                case ContentType.PPTX:
                    converter = new PowerpointConverter();
                    break;
                default:
                    break;
            }
            return converter;
        }
    }
}
