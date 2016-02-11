﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessDao = Microsoft.Office.Interop.Access.Dao;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeFileProperties.FileAccessors.Dao
{
    /// <summary>
    /// Class for using Microsoft Access databases.
    /// </summary>
    public class DaoFile : FileBase<AccessDao.Database>
    {

        #region Fields

        /// <summary>
        /// Database engine
        /// </summary>
        private readonly AccessDao.DBEngine _dbEngine;

        /// <summary>
        /// Database workspace
        /// </summary>
        private readonly AccessDao.Workspace _dbWorkspace;

        #endregion Fields

        #region Constructors

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

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Author"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
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
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Company"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
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
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                DateTime? createdTimeUtc = null;

                // Try to get time from SummaryInfo.
                try
                {
                    createdTimeUtc = DateTime.Parse(this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["DateCreated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();                    
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
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
                        if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                        {
                            return null;
                        }
                        else
                        {
                            // Rethrow the exception.
                            throw ce;
                        }
                    }
                }

                return createdTimeUtc;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, string> CustomProperties
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                try
                {
                    if (this.File.Containers["Databases"].Documents["UserDefined"].Properties == null)
                    {
                        return new Dictionary<string, string>();
                    }
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                var customProperties = new Dictionary<string, string>();

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
                                    customProperties.Add(property.Name.ToString(), property.Value.ToString());
                                    break;
                            }
                        }
                        catch (COMException ce)
                        {
                            // If conversion problem or value doesn't exist, return null
                            if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
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
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                return customProperties;
            }
        }

        /// <summary>
        /// Type of file.
        /// </summary>
        public override FileTypeEnum FileType
        {
            get
            {
                return FileTypeEnum.MicrosoftAccess;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedTimeUtc
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

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
                            try {
                                propertyTimeUtc = DateTime.Parse(document.Properties["LastUpdated"].Value.ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();
                            }
                            catch
                            { }

                            // Compare existing oldest and this property's times.
                            if (propertyTimeUtc.HasValue)
                            {
                                // Set modified time if it is still null.
                                modifiedTimeUtc = modifiedTimeUtc ?? propertyTimeUtc;

                                modifiedTimeUtc = (modifiedTimeUtc.Value < propertyTimeUtc.Value) ? propertyTimeUtc : modifiedTimeUtc;
                            }

                        }
                        catch (COMException ce)
                        {
                            // If conversion problem or value doesn't exist, return null
                            if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                            {
                                return null;
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
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                return modifiedTimeUtc;
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                try
                {
                    return this.File.Containers["Databases"].Documents["SummaryInfo"].Properties["Title"].Value.ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem or value doesn't exist, return null
                    if (((uint)ce.ErrorCode == 0x800A0D1E) || ((uint)ce.ErrorCode == 0x800A0CC6) || ((uint)ce.ErrorCode == 0x800A0CC1))
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public override void CloseFile()
        {
            // Mark file as closed.
            this.IsOpen = false;

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
        public override void OpenFile()
        {
            if (this._dbWorkspace == null)
            {
                throw new InvalidOperationException("Workspace is not ready.");
            }

            // Open file.
            this.File = this._dbWorkspace.OpenDatabase(this.Filename, false, true, "");

            // Mark file as open.
            this.IsOpen = true;
        }

        #endregion Methods

    }
}
