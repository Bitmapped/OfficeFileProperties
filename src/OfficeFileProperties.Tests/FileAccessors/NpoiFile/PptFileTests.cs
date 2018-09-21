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
        #region Methods

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
        public void PptGetCommentsTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

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

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 15, 41, 11, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptOpenAndCloseFileTest()
        {
            var file = new NpoiFile(@"..\..\SampleFiles\Test.Ppt");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}