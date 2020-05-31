using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Office2Pdf.Tests
{
    [TestClass]
    public class ExcelConverterTest
    {
        [TestMethod]
        public void Should_Convert_Xls()
        {
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.xls");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testxls.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Xlsx()
        {
            IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.XLSX);

            var sourcePath = Path.Combine(@"c:\Test.xslx");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testxslx.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_MoreTimes()
        {
            var index = 0;
            while (index < 10) {
                IConverter conveter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

                var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.xls");
                var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", $"Testxls_{index}.pdf");

                conveter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
