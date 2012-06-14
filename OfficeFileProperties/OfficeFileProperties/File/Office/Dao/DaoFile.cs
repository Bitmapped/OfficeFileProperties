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
    class DaoFile : IFile, IDisposable
    {
        // Define private variables.
        private string filename;
        private AccessDao.Database file;
        private OfficeFileProperties fileProperties;
        private bool disposed = false;

        // Instantiate shared Access Dao database objects.
        private AccessDao.DBEngine dbEngine;
        private AccessDao.Workspace dbWorkspace;

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

            // Instantiate Dao engine and workspace.
            this.dbEngine = new AccessDao.DBEngine();
            this.dbWorkspace = dbEngine.CreateWorkspace("", "admin", "", AccessDao.WorkspaceTypeEnum.dbUseJet);
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

            // Load file.
            this.file = this.dbWorkspace.OpenDatabase(filename, false, true, "");

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
            this.fileProperties.fileType = FileTypeEnum.MicrosoftAccess;
            
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
                        // If error code 0x800A0D1E or 0x800a0cc1, known missing database components.  Throw away.
                        if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800a0cc1))
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
                // If error code 0x800A0D1E or 0x800a0cc1, known missing database components.  Throw away.
                if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800a0cc1))
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

        /// <summary>
        /// Dispose of file access.
        /// </summary>
        /// <param name="disposing">If connection should be disposed.</param>
        protected void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    // Close file if not already done.
                    if (this.file != null)
                    {
                        // Close file.
                        this.CloseFile();
                    }

                    // Set context handlers to null.
                    this.dbEngine = null;
                    this.dbWorkspace = null;
                }
            }
            this.disposed = true;
        }

        /// <summary>
        /// Dispose of file access.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }
    }
}
