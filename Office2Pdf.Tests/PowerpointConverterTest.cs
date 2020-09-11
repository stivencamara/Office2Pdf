using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Office2Pdf.Tests
{
    [TestClass]
    public class PowerpointConverterTest
    {
        [TestMethod]
        public void Should_Convert_Ppt()
        {
            var conveter = new DocumentConverterFactory().GetConverter(ContentType.PPT);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.ppt");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testppt.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Pptx()
        {
            var converter = new DocumentConverterFactory().GetConverter(ContentType.PPTX);

            var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.pptx");
            var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", "Testpptx.pdf");

            converter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_MoreTimes()
        {
            var index = 0;
            while (index < 50) 
            {
                var converter = new DocumentConverterFactory().GetConverter(ContentType.PPT);

                var sourcePath = Path.Combine(Environment.CurrentDirectory, "docs", "Test.ppt");
                var targetPath = Path.Combine(Environment.CurrentDirectory, "docs", $"Testdoc_{index}.pdf");

                converter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
