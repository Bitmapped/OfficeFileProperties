using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Dao.Tests
{
    [TestClass()]
    public class MdbFileTests
    {
        #region Methods

        [TestMethod()]
        public void MdbGetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Mdb");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Mdb");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetCreatedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Mdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 15, 24, 25, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod]
        public void MdbGetCustomPropertiesTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            var expectedValue = new Dictionary<string, object>() { { "Test1", "Test" }, { "Test2", 1 } };

            CollectionAssert.AreEqual((ICollection)expectedValue, (ICollection)file.CustomProperties);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetModifiedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Mdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 16, 10, 56, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbOpenAndCloseFileTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Mdb");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}