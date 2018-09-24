using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace OfficeFileProperties.FileAccessors.Dao.Tests
{
    [TestClass()]
    public class MdbFileTests
    {
        #region Methods

        [TestMethod()]
        public void MdbGetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbSetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.mdb");
            var testValue = $"Test Author {DateTime.Now}";

            file.OpenFile(true);
            file.Author = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Author);
            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbSetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.mdb");
            var testValue = $"Test Company {DateTime.Now}";

            file.OpenFile(true);
            file.Company = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Company);
            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetTitleTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual("Test Title", file.Title);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbSetTitleTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.mdb");
            var testValue = $"Test Title {DateTime.Now}";

            file.OpenFile(true);
            file.Title = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Title);
            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetCommentsTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbSetCommentsTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.mdb");
            var testValue = $"Test Comments {DateTime.Now}";

            file.OpenFile(true);
            file.Comments = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Comments);
            file.CloseFile();
        }

        [TestMethod()]
        public void MdbGetCreatedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 15, 24, 25, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }


        [TestMethod()]
        public void MdbGetModifiedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 14, 35, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void MdbOpenAndCloseFileTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.mdb");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}