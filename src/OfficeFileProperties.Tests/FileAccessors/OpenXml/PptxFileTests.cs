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
    public class PptxFileTests
    {
        #region Methods

        [TestMethod()]
        public void PptxGetAuthorTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetAuthorTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = $"Test Author {DateTime.Now}";

            file.OpenFile(true);
            file.Author = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Author);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = $"Test Company {DateTime.Now}";

            file.OpenFile(true);
            file.Company = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Company);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetTitleTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual("Test Title", file.Title);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetTitleTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = $"Test Title {DateTime.Now}";

            file.OpenFile(true);
            file.Title = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Title);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCommentsTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetCommentsTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = $"Test Comments {DateTime.Now}";

            file.OpenFile(true);
            file.Comments = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.Comments);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 57, 59, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = DateTime.UtcNow.AddYears(1);

            file.OpenFile(true);
            file.CreatedTimeUtc = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.CreatedTimeUtc);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 15, 55, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxSetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\WriteTest.pptx");
            var testValue = DateTime.UtcNow.AddYears(5);

            file.OpenFile(true);
            file.ModifiedTimeUtc = testValue;
            file.CloseFile();

            file.OpenFile();
            Assert.AreEqual(testValue, file.ModifiedTimeUtc);
            file.CloseFile();
        }

        [TestMethod()]
        public void PptxOpenAndCloseFileTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.pptx");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}