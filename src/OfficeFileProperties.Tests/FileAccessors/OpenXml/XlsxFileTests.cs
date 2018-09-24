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
        #region Methods

        [TestMethod()]
        public void XlsxGetAuthorTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetAuthorTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = $"Test Author {DateTime.Now}";

            file.OpenFile(true);
            file.Author = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Author);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetCompanyTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetCompanyTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = $"Test Company {DateTime.Now}";

            file.OpenFile(true);
            file.Company = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Company);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetTitleTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual("Test Title", file.Title);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetTitleTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = $"Test Title {DateTime.Now}";

            file.OpenFile(true);
            file.Title = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Title);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetCommentsTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetCommentsTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = $"Test Comments {DateTime.Now}";

            file.OpenFile(true);
            file.Comments = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Comments);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetCreatedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 29, 26, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetCreatedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = DateTime.UtcNow.AddYears(1);

            file.OpenFile(true);
            file.CreatedTimeUtc = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.CreatedTimeUtc);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxGetModifiedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 15, 13, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxSetModifiedTimeUtcTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\WriteTest.xlsx");
            var testValue = DateTime.UtcNow.AddYears(5);

            file.OpenFile(true);
            file.ModifiedTimeUtc = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.ModifiedTimeUtc);
            file.CloseFile();
        }

        [TestMethod()]
        public void XlsxOpenAndCloseFileTest()
        {
            var file = new XlsxFile(@"..\..\SampleFiles\Test.xlsx");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}