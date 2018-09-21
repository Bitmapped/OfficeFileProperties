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
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCompanyTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCommentsTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetCreatedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 57, 59, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxGetModifiedTimeUtcTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 15, 55, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void PptxOpenAndCloseFileTest()
        {
            var file = new PptxFile(@"..\..\SampleFiles\Test.Pptx");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}