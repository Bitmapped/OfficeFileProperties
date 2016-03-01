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

            // Test property.
            if (file.Author != "Test Author")
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            // Test property.
            if (file.Company != "Test Company")
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            // Test property.
            if (file.CreatedTimeUtc != new DateTime(2016, 3, 1, 3, 57, 59, DateTimeKind.Utc))
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            // Test property.s
            if (file.ModifiedTimeUtc != new DateTime(2016, 3, 1, 3, 58, 29, DateTimeKind.Utc))
            {
                Assert.Fail();
            }

            file.CloseFile();
        }
    }
}