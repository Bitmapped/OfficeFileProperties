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
    public class DocxFileTests
    {
        [TestMethod()]
        public void DocxOpenAndCloseFileTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetAuthorTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            // Test property.
            if (file.Author != "Test Author")
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetCompanyTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            // Test property.
            if (file.Company != "Test Company")
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetCreatedTimeUtcTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            // Test property.
            if (file.CreatedTimeUtc != new DateTime(2016, 3, 1, 3, 53, 0, DateTimeKind.Utc))
            {
                Assert.Fail();
            }

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetModifiedTimeUtcTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            // Test property.s
            if (file.ModifiedTimeUtc != new DateTime(2016, 3, 1, 3, 56, 0, DateTimeKind.Utc))
            {
                Assert.Fail();
            }

            file.CloseFile();
        }
    }
}