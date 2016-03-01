using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Npoi.Tests
{
    [TestClass()]
    public class DocFileTests
    {
        #region Methods

        [TestMethod()]
        public void DocGetAuthorTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetCompanyTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetCreatedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 54, 0, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocGetModifiedTimeUtcTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 55, 0, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocOpenAndCloseFileTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Doc");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}