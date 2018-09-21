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
    public class DocxFileTests
    {
        #region Methods

        [TestMethod()]
        public void DocxGetAuthorTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            Assert.AreEqual("Test Author", file.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetCompanyTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            Assert.AreEqual("Test Company", file.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetCommentsTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            Assert.AreEqual("Test Comments", file.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetCreatedTimeUtcTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 53, 0, DateTimeKind.Utc), file.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxGetModifiedTimeUtcTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 11, 0, DateTimeKind.Utc), file.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void DocxOpenAndCloseFileTest()
        {
            var file = new DocxFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}