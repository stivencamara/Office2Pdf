using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Office2Pdf.Tests
{
    [TestClass]
    public class ExcelConverterTest
    {
        private string directoty = Path.Combine(Environment.CurrentDirectory, "Docs");

        [TestMethod]
        public void Should_Convert_Xls()
        {
            var conveter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

            var sourcePath = Path.Combine(directoty, "Test.xls");
            var targetPath = Path.Combine(directoty, "Testxls.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Xlsx()
        {
            var conveter = new DocumentConverterFactory().GetConverter(ContentType.XLSX);

            var sourcePath = Path.Combine(directoty, "Test.xlsx");
            var targetPath = Path.Combine(directoty, "Testxslx.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_MoreTimes()
        {
            var index = 0;
            while (index < 10) {
                var converter = new DocumentConverterFactory().GetConverter(ContentType.XLS);

                var sourcePath = Path.Combine(directoty, "Test.xls");
                var targetPath = Path.Combine(directoty, $"Testxls_{index}.pdf");

                converter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
