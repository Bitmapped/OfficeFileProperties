using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Npoi.Tests
{
    [TestClass()]
    public class PptFileTests
    {
        [TestMethod()]
        public void PptOpenAndCloseFileTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void PptGetAuthorTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptGetCompanyTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptGetCreatedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 57, 59, 474, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptGetModifiedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 58, 35, 899, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }
    }
}