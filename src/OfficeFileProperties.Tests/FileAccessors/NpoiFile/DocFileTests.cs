using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace OfficeFileProperties.FileAccessors.Npoi.Tests
{
    [TestClass()]
    public class DocFileTests
    {
        [TestMethod()]
        public void DocOpenAndCloseFileTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetAuthorTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            //throw new Exception(file.FileProperties.DocumentSummaryInformationText);

            Assert.AreEqual(file.Author, "Test Author");

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetCompanyTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual(file.Company, "Test Company");

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetCreatedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual(file.CreatedTimeUtc, new DateTime(2016, 3, 1, 3, 54, 0, DateTimeKind.Utc));

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetModifiedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual(file.ModifiedTimeUtc, new DateTime(2016, 3, 1, 3, 55, 0, DateTimeKind.Utc));

            file.CloseFile();
        }
    }
}