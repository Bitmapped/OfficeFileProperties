using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using AccessDao = Microsoft.Office.Interop.Access.Dao;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeFileProperties.File.Office.Dao
{
    /// <summary>
    /// Obtain properties of Microsoft Access databases.
    /// </summary>
    class DaoFile : IFile
    {
        // Define private variables.
        private string filename;
        private AccessDao.Database file;
        private OfficeFileProperties fileProperties;

        /// <summary>
        /// Access file properties.
        /// </summary>
        public IFileProperties FileProperties
        {
            get
            {
                // Get file if loaded.
                if ((this.fileProperties != null) && (this.fileProperties.fileLoaded))
                {
                    return this.fileProperties;
                }
                else
                {
                    // File has not been loaded.
                    throw new InvalidOperationException("No file has been loaded.");
                }
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public DaoFile()
        {
            ClearProperties();
        }

        /// <summary>
        /// Constructor, also loads specified file.
        /// </summary>
        /// <param name="filename">File to load.</param>
        public DaoFile(string filename)
            : this()
        {
            // Load specified file.
            this.LoadFile(filename);
        }

        /// <summary>
        /// Forget about loaded file.
        /// </summary>
        private void CloseFile()
        {
            // Close file.
            this.file.Close();

            // Clear file object.
            this.file = null;
        }

        /// <summary>
        /// Clears values of loaded properties.
        /// </summary>
        private void ClearProperties()
        {
            // Clear loaded file, properties object.
            this.filename = null;
            this.fileProperties = null;
        }

        /// <summary>
        /// Loads requested file, saves its properties.
        /// </summary>
        /// <param name="filename"></param>
        public void LoadFile(string filename)
        {
            // Clear loaded properties.
            ClearProperties();

            // Store filename.
            this.filename = filename;

            // Instantiate Dao engine and workspace.
            var dbEngine = new AccessDao.DBEngine();
            var dbWorkspace = dbEngine.CreateWorkspace("", "admin", "", AccessDao.WorkspaceTypeEnum.dbUseJet);

            // Load file.
            this.file = dbWorkspace.OpenDatabase(filename, false, true, "");

            // Loads file properties.
            LoadProperties();

            // Since file cannot be written to, close it immediately.
            CloseFile();
        }

        /// <summary>
        /// Load values from file into private variables.
        /// </summary>
        private void LoadProperties()
        {
            // If file hasn't been loaded, throw exception.
            if (this.file == null)
            {
                throw new InvalidOperationException("No file is currently loaded.");
            }

            // Create new file properties object.
            this.fileProperties = new OfficeFileProperties();

            // filename
            this.fileProperties.filename = this.filename;

            // filetype
            this.fileProperties.fileType = OfficeFileProperties.FileTypeEnum.MicrosoftAccess;
            
            // createdTimeUtc
            // Try getting actual time, first through SummaryInfo.
            try
            {
                this.fileProperties.createdTimeUtc = DateTime.Parse(this.file.Containers["Databases"].Documents["SummaryInfo"].Properties["DateCreated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
            }
            catch
            {
                // Try getting actual time through MSysDb.
                try
                {
                    this.fileProperties.createdTimeUtc = DateTime.Parse(this.file.Containers["Databases"].Documents["MSysDb"].Properties["DateCreated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                }
                catch
                {
                    this.fileProperties.createdTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                }
            }

            // modifiedTimeUtc
            // Try getting actual time, otherwise return dummy value if it fails.

                this.fileProperties.modifiedTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);

                // Loop through all document items to find newer time.
                try
                {
                    DateTime updatedTime;
                    foreach (AccessDao.Container container in this.file.Containers)
                    {
                        foreach (AccessDao.Document document in container.Documents)
                        {
                            // Get time of object.
                            updatedTime = DateTime.Parse(document.Properties["LastUpdated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();

                            // Compare time to already-saved time.
                            if (updatedTime > this.fileProperties.modifiedTimeUtc)
                            {
                                // New time is more recent.  Save it.
                                this.fileProperties.modifiedTimeUtc = updatedTime;
                            }
                        }
                    }
                }
                catch { }


            // author
            // Try obtaining, returning null if not available.
            try
            {
                this.fileProperties.author = this.file.Containers["Databases"].Documents["SummaryInfo"].Properties["Author"].Value.ToString();
            }
            catch
            {
                this.fileProperties.author = null;
            }

            // title
            // Try obtaining, returning null if not available.
            try
            {
                this.fileProperties.title = this.file.Containers["Databases"].Documents["SummaryInfo"].Properties["Title"].Value.ToString();
            }
            catch
            {
                this.fileProperties.title = null;
            }

            // company
            // Try obtaining, returning null if not available.
            try
            {
                this.fileProperties.company = this.file.Containers["Databases"].Documents["SummaryInfo"].Properties["Company"].Value.ToString();
            }
            catch
            {
                this.fileProperties.company = null;
            }

            // Load custom properties.
            // Try obtaining.
            try
            {
                foreach (AccessDao.Property cp in this.file.Containers["Databases"].Documents["UserDefined"].Properties)
                {
                    // Include in try-catch block to cover over bad name exceptions.
                    try
                    {
                        switch (cp.Name)
                        {
                            case "Name":
                            case "Owner":
                            case "UserName":
                            case "Container":
                            case "DateCreated":
                            case "LastUpdated":
                                // Default properties.  Do nothing.
                                break;

                            default:
                                // Record property.
                                this.fileProperties.customProperties.Add(cp.Name.ToString(), cp.Value.ToString());
                                break;
                        }
                    }
                    catch (COMException ce)
                    {
                        // If error code 0x800A0D1E, known error - throw away.
                        if ((uint)ce.ErrorCode == 0x800A0D1E)
                        {
                            // Do nothing.
                        }
                        else
                        {
                            // Rethrow the exception.
                            throw ce;
                        }
                    }
                }
            }
            catch (COMException ce)
            {
                // If error code 0x800A0D1E, known error - throw away.
                if ((uint)ce.ErrorCode == 0x800A0D1E)
                {
                    // Do nothing.
                }
                else
                {
                    // Rethrow the exception.
                    throw ce;
                }
            }

            // Mark properties as loaded.
            this.fileProperties.fileLoaded = true;
        }

    }
}
