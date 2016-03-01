using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace OfficeFileProperties.FileAccessors.OpenXml.Tests
{
    [TestClass()]
    public class PptxFileTests
    {
        [TestMethod()]
        public void PptxOpenAndCloseFileTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetAuthorTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 57, 59, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 58, 29, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }
    }
}