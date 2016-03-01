using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Dao.Tests
{
    [TestClass()]
    public class AccdbFileTests
    {
        [TestMethod()]
        public void AccdbOpenAndCloseFileTest()
        {
            var file = new DaoFile(@"..\..\SampleFiles\Test.Accdb");
            file.OpenFile();
            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetAuthorTest()
        {
            var file = new DaoFile(@"..\..\SampleFiles\Test.Accdb");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetCompanyTest()
        {
            var file = new DaoFile(@"..\..\SampleFiles\Test.Accdb");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetCreatedTimeUtcTest()
        {
            var file = new DaoFile(@"..\..\SampleFiles\Test.Accdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 15, 24, 25, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetModifiedTimeUtcTest()
        {
            var file = new DaoFile(@"..\..\SampleFiles\Test.Accdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 15, 27, 23, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }
    }
}