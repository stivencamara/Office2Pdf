using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Office2Pdf.Tests
{
    [TestClass]
    public class WordConverterTest
    {
        private string directoty = Path.Combine(Environment.CurrentDirectory, "Docs");

        [TestMethod]
        public void Should_Convert_Docx()
        {
            var converter = new DocumentConverterFactory().GetConverter(ContentType.DOCX);

            var sourcePath = Path.Combine(directoty, "Test.docx");
            var targetPath = Path.Combine(directoty, "Testdocx.pdf");

            converter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Doc()
        { 
            var converter = new DocumentConverterFactory().GetConverter(ContentType.DOC);

            var sourcePath = Path.Combine(directoty, "Test.doc");
            var targetPath = Path.Combine(directoty, "Testdoc.pdf");

            converter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_MoreTimes()
        {
            var index = 0;
            while (index < 50) 
            {
                var converter = new DocumentConverterFactory().GetConverter(ContentType.DOC);

                var sourcePath = Path.Combine(directoty, "Test.doc");
                var targetPath = Path.Combine(directoty, $"Testdoc_{index}.pdf");

                converter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
