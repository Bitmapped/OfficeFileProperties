using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Npoi.Tests
{
    [TestClass()]
    public class XlsFileTests
    {
        [TestMethod()]
        public void XlsOpenAndCloseFileTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Xls");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsGetAuthorTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Xls");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsGetCompanyTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Xls");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsGetCreatedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Xls");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 29, 26, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsGetModifiedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Xls");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 30, 47, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }
    }
}