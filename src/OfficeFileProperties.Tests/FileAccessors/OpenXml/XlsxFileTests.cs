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
    public class XlsxFileTests
    {
        [TestMethod()]
        public void XlsxOpenAndCloseFileTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetAuthorTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(file.Author, "Test Author");

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetCompanyTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(file.Company, "Test Company");

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetCreatedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(file.CreatedTimeUtc, new DateTime(2016, 3, 1, 3, 29, 26, DateTimeKind.Utc));

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetModifiedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(file.ModifiedTimeUtc, new DateTime(2016, 3, 1, 3, 30, 33, DateTimeKind.Utc));

            file.CloseFile();
        }
    }
}