using System;
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

        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public DaoFile(string filename) : base(filename)
        {
            // Initialize DAO engine and workspace.
            this.DbEngine = new AccessDao.DBEngine();
            this.DbWorkspace = this.DbEngine.CreateWorkspace("", "admin", "", AccessDao.WorkspaceTypeEnum.dbUseJet);
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
                try
                {
                    return this.FileAccessor.Containers["Databases"].Documents["SummaryInfo"].Properties["Author"].ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
                try
                {
                    return this.FileAccessor.Containers["Databases"].Documents["SummaryInfo"].Properties["Company"].ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
        public override DateTime? CreatedDateUtc
        {
            get
            {
                DateTime? createdTimeUtc = null;
                try
                {

                    createdTimeUtc = DateTime.Parse(this.FileAccessor.Containers["Databases"].Documents["MSysDb"].Properties["DateCreated"].ToString(), new CultureInfo("en-US"), DateTimeStyles.AssumeLocal).ToUniversalTime();                    
                }
                catch (COMException ce)
                {
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
                    {
                        return null;
                    }
                    else
                    {
                        // Rethrow the exception.
                        throw ce;
                    }
                }

                return createdTimeUtc;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override Dictionary<string, string> CustomProperties
        {
            get
            {
                try
                {
                    if (this.FileAccessor.Containers["Databases"].Documents["UserDefined"].Properties == null)
                    {
                        return new Dictionary<string, string>();
                    }
                }
                catch (COMException ce)
                {
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
                    foreach (AccessDao.Property property in this.FileAccessor.Containers["Databases"].Documents["UserDefined"].Properties)
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
                            // If conversion problem, throw away exception.
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
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
        public override DateTime? ModifiedDateUtc
        {
            get
            {
                DateTime? modifiedTimeUtc = null;

                // Iterate through database items to find last modified time.
                try
                {
                    foreach (AccessDao.Document document in this.FileAccessor.Containers["Databases"].Documents)
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

                                modifiedTimeUtc = (modifiedTimeUtc.Value < propertyTimeUtc.Value) ? modifiedTimeUtc : propertyTimeUtc;
                            }

                        }
                        catch (COMException ce)
                        {
                            // If conversion problem, throw away exception.
                            if ((uint)ce.ErrorCode == 0x800A0D1E)
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
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
                try
                {
                    return this.FileAccessor.Containers["Databases"].Documents["SummaryInfo"].Properties["Title"].ToString();
                }
                catch (COMException ce)
                {
                    // If conversion problem, throw away exception.
                    if ((uint)ce.ErrorCode == 0x800A0D1E)
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
        /// Database engine
        /// </summary>
        private AccessDao.DBEngine DbEngine { get; set; }

        /// <summary>
        /// Database workspace
        /// </summary>
        private AccessDao.Workspace DbWorkspace { get; set; }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public override void CloseFile()
        {
            // Mark file as closed.
            this.IsOpen = false;

            // Close file.
            this.FileAccessor.Close();

            // Clear file object.
            this.FileAccessor = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public override void OpenFile()
        {
            // Open file.
            this.FileAccessor = this.DbWorkspace.OpenDatabase(this.Filename, false, true, "");

            // Mark file as open.
            this.IsOpen = true;
        }

        #endregion Methods
    }
}
