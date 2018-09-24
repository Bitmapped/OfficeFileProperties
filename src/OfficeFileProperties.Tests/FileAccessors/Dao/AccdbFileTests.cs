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
    public class AccdbFileTests
    {
        #region Methods

        [TestMethod()]
        public void AccdbGetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbSetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.accdb");
            var testValue = $"Test Author {DateTime.Now}";

            file.OpenFile(true);
            file.Author = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Author);
            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbSetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.accdb");
            var testValue = $"Test Company {DateTime.Now}";

            file.OpenFile(true);
            file.Company = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Company);
            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetTitleTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual("Test Title", file.Title);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbSetTitleTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.accdb");
            var testValue = $"Test Title {DateTime.Now}";

            file.OpenFile(true);
            file.Title = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Title);
            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetCommentsTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbSetCommentsTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\WriteTest.accdb");
            var testValue = $"Test Comments {DateTime.Now}";

            file.OpenFile(true);
            file.Comments = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Comments);
            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbGetCreatedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 15, 24, 25, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }


        [TestMethod()]
        public void AccdbGetModifiedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 16, 02, 33, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void AccdbOpenAndCloseFileTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.accdb");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}