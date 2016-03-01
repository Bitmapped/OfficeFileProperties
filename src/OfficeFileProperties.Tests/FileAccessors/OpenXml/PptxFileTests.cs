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

            Assert.AreEqual(file.Author, "Test Author");

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(file.Company, "Test Company");

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(file.CreatedTimeUtc, new DateTime(2016, 3, 1, 3, 57, 59, DateTimeKind.Utc));

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(file.ModifiedTimeUtc, new DateTime(2016, 3, 1, 3, 58, 29, DateTimeKind.Utc));

            file.CloseFile();
        }
    }
}