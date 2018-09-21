using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeFileProperties.Tests.FileAccessors
{
    [TestClass]
    public class OfficeFilePropertiesTests
    {
        #region Methods

        [TestMethod()]
        public void FilePropertiesDocxGetAuthorTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            var fileProperties = file.GetFileProperties();

            Assert.AreEqual("Test Author", fileProperties.Author);

            file.CloseFile();
        }

        [TestMethod()]
        public void FilePropertiesDocxGetCompanyTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            var fileProperties = file.GetFileProperties();

            Assert.AreEqual("Test Company", fileProperties.Company);

            file.CloseFile();
        }

        [TestMethod()]
        public void FilePropertiesDocxGetCommentsTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            var fileProperties = file.GetFileProperties();

            Assert.AreEqual("Test Comments", fileProperties.Comments);

            file.CloseFile();
        }

        [TestMethod()]
        public void FilePropertiesDocxGetCreatedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            var fileProperties = file.GetFileProperties();

            Assert.AreEqual(new DateTime(2016, 3, 1, 3, 53, 0, DateTimeKind.Utc), fileProperties.CreatedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void FilePropertiesDocxGetModifiedTimeUtcTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            var fileProperties = file.GetFileProperties();

            Assert.AreEqual(new DateTime(2018, 9, 21, 15, 11, 0, DateTimeKind.Utc), fileProperties.ModifiedTimeUtc);

            file.CloseFile();
        }

        [TestMethod()]
        public void FilePropertiesDocxOpenAndCloseFileTest()
        {
            var file = new OfficeFile(@"..\..\SampleFiles\Test.Docx");
            file.OpenFile();
            file.CloseFile();
        }

        #endregion Methods
    }
}
