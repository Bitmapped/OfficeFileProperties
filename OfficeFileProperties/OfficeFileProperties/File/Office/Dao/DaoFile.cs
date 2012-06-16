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

            // Load properties.
            AccessDao.Document summaryInfo = null, userDefined = null;
            try
            {
                // Extract documents in which we have interest.  Note most recent updated time.
                this.fileProperties.modifiedTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                DateTime updatedTime;
                DateTime? mSysDbTime = null;
                foreach (AccessDao.Document document in this.file.Containers["Databases"].Documents)
                {
                    switch (document.Name)
                    {
                        case "SummaryInfo":
                            summaryInfo = document;
                            break;

                        case "UserDefined":
                            userDefined = document;
                            break;

                        default:
                            break;
                    }

                    // Compare edit times.
                    foreach (AccessDao.Property property in document.Properties)
                    {
                        // Look for name of LastUpdated.
                        if (property.Name == "LastUpdated")
                        {
                            // Get time of object.
                            updatedTime = DateTime.Parse(property.Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();

                            // Compare time to already-saved time.
                            if (updatedTime > this.fileProperties.modifiedTimeUtc)
                            {
                                // New time is more recent.  Save it.
                                this.fileProperties.modifiedTimeUtc = updatedTime;
                            }
                        }
                        else
                        {
                            // Set aside mSysDb.DateCreated for potential later use if it exists.
                            if ((document.Name == "MSysDb") && (property.Name == "DateCreated"))
                            {
                                mSysDbTime = DateTime.Parse(property.Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                            }

                        }
                    }
                }

                // Process properties from summaryInfo.
                if (summaryInfo != null)
                {
                    foreach (AccessDao.Property property in summaryInfo.Properties)
                    {
                        switch (property.Name)
                        {
                            case "DateCreated":
                                this.fileProperties.createdTimeUtc = DateTime.Parse(property.Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                                break;

                            case "Author":
                                this.fileProperties.author = property.Value.ToString();
                                break;

                            case "Title":
                                this.fileProperties.title = property.Value.ToString();
                                break;

                            case "Company":
                                this.fileProperties.company = property.Value.ToString();
                                break;

                            default:
                                break;
                        }
                    }
                }

                // If summaryInfo didn't provide created time, try getting from mSysDb.
                if (this.fileProperties.createdTimeUtc == null)
                {
                    if (mSysDbTime != null)
                    {
                        this.fileProperties.createdTimeUtc = (DateTime)mSysDbTime;
                    }
                    else
                    {
                        // Use generic date of 1/1/0001.
                        this.fileProperties.createdTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                    }
                }

                // Process user-defined properties.
                if (userDefined != null)
                {
                    foreach (AccessDao.Property property in userDefined.Properties)
                    {
                        switch (property.Name)
                        {
                            case "Name":
                            case "Owner":
                            case "UserName":
                            case "Container":
                            case "DateCreated":
                            case "LastUpdated":
                                // Default properties.  Do nothing.
                                break;

                            case "Permissions":
                            case "AllPermissions":
                                // Can't handle this property.  Do nothing.
                                break;

                            default:
                                // Record property.
                                this.fileProperties.customProperties.Add(property.Name.ToString(), property.Value.ToString());
                                break;
                        }
                    }
                }

            }
            catch (Exception e)
            {
                throw e;
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
