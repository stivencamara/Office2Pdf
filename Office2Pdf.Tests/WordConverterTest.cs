using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Convert2Office.Tests
{
    [TestClass]
    public class WordConverterTest
    {
        [TestMethod]
        public void Should_Convert_Docx()
        {
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.DOCX);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.docx");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testdocx.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Doc()
        {
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.DOC);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.doc");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testdoc.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_MoreTimes()
        {
            var index = 0;
            while (index < 50) {
                IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.DOC);

                var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.doc");
                var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", $"Testdoc_{index}.pdf");

                conveter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
