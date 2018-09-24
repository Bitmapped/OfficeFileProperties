using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using AccessDao = Microsoft.Office.Interop.Access.Dao;

namespace OfficeFileProperties.FileAccessors.Dao
{
    /// <summary>
    /// Class for using Microsoft Access databases.
    /// </summary>
    public class DaoFile : FileBase<AccessDao.Database>
    {
        /// <summary>
        /// Database engine
        /// </summary>
        private readonly AccessDao.DBEngine _dbEngine;

        /// <summary>
        /// Database workspace
        /// </summary>
        private readonly AccessDao.Workspace _dbWorkspace;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public DaoFile(string filename) : base(filename)
        {
            // Initialize DAO engine and workspace.
            this._dbEngine = new AccessDao.DBEngine();
            this._dbWorkspace = this._dbEngine.CreateWorkspace("", "admin", "", AccessDao.WorkspaceTypeEnum.dbUseJet);
        }

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Author"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Try to delete existing property.
                try
                {
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Delete("Author");
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        // Do nothing.
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                // Try to set new property.
                try
                {
                    var prop = this.File.CreateProperty("Author", AccessDao.DataTypeEnum.dbText, value, true);
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Append(prop);
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
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

        /// <summary>
        /// Comments (description)
        /// </summary>
        public override string Comments
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Comments"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Try to delete existing property.
                try
                {
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Delete("Comments");
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        // Do nothing.
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                // Try to set new property.
                try
                {
                    var prop = this.File.CreateProperty("Comments", AccessDao.DataTypeEnum.dbText, value, true);
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Append(prop);
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
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

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Company"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Try to delete existing property.
                try
                {
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Delete("Company");
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        // Do nothing.
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                // Try to set new property.
                try
                {
                    var prop = this.File.CreateProperty("Company", AccessDao.DataTypeEnum.dbText, value, true);
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Append(prop);
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
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

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                DateTime? createdTimeUtc = null;

                // Try to get time from SummaryInfo.
                try
                {
                    createdTimeUtc = DateTime.Parse(this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["DateCreated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }

                // If time could not be obtained from SummaryInfo, try MSysDB.
                if (!createdTimeUtc.HasValue)
                {
                    try
                    {
                        createdTimeUtc = DateTime.Parse(this.File.Containers["Databases"].Documents["MSysDb"].Properties["DateCreated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                    }
                    catch (COMException ce)
                    {
                        // If conversion problem or value doesn't exist, return null
                        if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                        {
                            return null;
                        }

                        // Rethrow the exception.
                        throw ce;
                    }
                }

                return createdTimeUtc;
            }
            set => throw new InvalidOperationException();
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, object> CustomProperties
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                try
                {
                    if (this.File.Containers["Databases"].Documents["UserDefined"].Properties == null)
                    {
                        return new Dictionary<string, object>();
                    }
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }

                var customProperties = new Dictionary<string, object>();

                try
                {
                    foreach (AccessDao.Property property in this.File.Containers["Databases"].Documents["UserDefined"].Properties)
                    {
                        try
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
                                    customProperties.Add(property.Name, property.Value);
                                    break;
                            }
                        }
                        catch (COMException ce)
                        {
                            // If conversion problem or value doesn't exist, return null
                            if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
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
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }

                return customProperties;
            }
        }

        /// <summary>
        /// Type of file.
        /// </summary>
        public override FileTypeEnum FileType => FileTypeEnum.MicrosoftAccess;

        /// <summary>
        /// Indicator if the file is open.
        /// </summary>
        public override bool IsOpen => this.File != null;

        /// <summary>
        /// Indicator if the file is readable.
        /// </summary>
        public override bool IsReadable => this.File != null;

        /// <summary>
        /// Indicator if the file is writable.
        /// </summary>
        public override bool IsWritable => this.File.Updatable;

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedTimeUtc
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                DateTime? modifiedTimeUtc = null;

                // Iterate through database items to find last modified time.
                try
                {
                    foreach (AccessDao.Document document in this.File.Containers["Databases"].Documents)
                    {
                        try
                        {
                            // Determine local time.
                            DateTime? propertyTimeUtc = null;
                            try
                            {
                                propertyTimeUtc = DateTime.Parse(document.Properties["LastUpdated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                            }
                            catch
                            { }

                            // Compare existing oldest and this property's times.
                            if (propertyTimeUtc.HasValue)
                            {
                                // Set modified time if it is still null.
                                modifiedTimeUtc = modifiedTimeUtc ?? propertyTimeUtc;

                                modifiedTimeUtc = modifiedTimeUtc.Value < propertyTimeUtc.Value ? propertyTimeUtc : modifiedTimeUtc;
                            }
                        }
                        catch (COMException ce)
                        {
                            // If conversion problem or value doesn't exist, return null
                            if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                            {
                                return null;
                            }

                            // Rethrow the exception.
                            throw ce;
                        }
                    }
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }

                return modifiedTimeUtc;
            }
            set => throw new InvalidOperationException();
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Title"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        return null;
                    }

                    // Rethrow the exception.
                    throw ce;
                }
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Try to delete existing property.
                try
                {
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Delete("Title");
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
                    {
                        // Do nothing.
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                // Try to set new property.
                try
                {
                    var prop = this.File.CreateProperty("Title", AccessDao.DataTypeEnum.dbText, value, true);
                    this.File.Containers["Databases"].Documents["SummaryInfo"].Properties.Append(prop);
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if ((uint) ce.ErrorCode == 0x800A0D1E || (uint) ce.ErrorCode == 0x800A0CC6 || (uint) ce.ErrorCode == 0x800A0CC1)
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

        /// <summary>
        /// Closes file.
        /// </summary>
        /// <param name="saveChanges"></param>
        public override void CloseFile(bool saveChanges = false)
        {
            // Close file if it still is accessible.
            if (this.File != null)
            {
                // Close file.
                this.File.Close();
            }

            // Clear file object.
            this.File = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable"></param>
        public override void OpenFile(bool writable = false)
        {
            if (this._dbWorkspace == null)
            {
                throw new InvalidOperationException("Workspace is not ready.");
            }

            // Open file.
            this.File = this._dbWorkspace.OpenDatabase(this.Filename, false, !writable, "");
        }
    }
}