using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;

namespace Office2Pdf.Tests
{
    [TestClass]
    public class PowerpointConverterTest
    {
        private string directoty = Path.Combine(Environment.CurrentDirectory, "Docs");

        [TestMethod]
        public void Should_Convert_Ppt()
        {
            var conveter = new DocumentConverterFactory().GetConverter(ContentType.PPT);

            var sourcePath = Path.Combine(directoty, "Test.ppt");
            var targetPath = Path.Combine(directoty, "Testppt.pdf");

            conveter.Convert(sourcePath, targetPath, false);

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void Should_Convert_Pptx()
        {
            var converter = new DocumentConverterFactory().GetConverter(ContentType.PPTX);

            var sourcePath = Path.Combine(directoty, "Test.pptx");
            var targetPath = Path.Combine(directoty, "Testpptx.pdf");

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

                var sourcePath = Path.Combine(directoty, "Test.ppt");
                var targetPath = Path.Combine(directoty, $"Testdoc_{index}.pdf");

                converter.Convert(sourcePath, targetPath, false);

                Thread.Sleep(200);

                index++;
            }

            Assert.IsTrue(true);
        }
    }
}
