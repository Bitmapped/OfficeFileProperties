using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OfficeFileProperties.File
{
    public class File : IFile
    {
        // Define private variables.
        private string filename;
        private IFile file;
        private IFileProperties fileProperties;
        private bool multifileMode = false, fileLoaded = false;

        // Store all accessors for multifile mode.
        private Office.Dao.DaoFile daoFile;
        private Office.Dso.DsoFile dsoFile;
        private Office.OpenXml.DocxFile docxFile;
        private Office.OpenXml.PptxFile pptxFile;
        private Office.OpenXml.XlsxFile xlsxFile;
        private Generic.GenericFile genericFile;

        /// <summary>
        /// Access file properties.
        /// </summary>
        public IFileProperties FileProperties
        {
            get
            {
                // Get file if loaded.
                if ((this.fileProperties != null) && (this.fileLoaded))
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
        /// <param name="multifileMode">If true, reuse file connection objects.</param>
        public File(bool multifileMode = false)
        {
            // Clear properties.
            ClearProperties();

            // Store multifile mode flag.
            this.multifileMode = multifileMode;

            // If multifile mode instantiated, create accessors.
            if (multifileMode)
            {
                daoFile = new Office.Dao.DaoFile();
                dsoFile = new Office.Dso.DsoFile();
                docxFile = new Office.OpenXml.DocxFile();
                pptxFile = new Office.OpenXml.PptxFile();
                xlsxFile = new Office.OpenXml.XlsxFile();
                genericFile = new Generic.GenericFile();
            }
        }

        /// <summary>
        /// Constructor to load specified file.
        /// </summary>
        /// <param name="filename">Filename to load.</param>
        /// <param name="multifileMode">If true, reuse file connection objects.</param>
        public File(string filename, bool multifileMode = false)
            : this(multifileMode)
        {
            this.LoadFile(filename);
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
        /// <param name="filename">Filename to open.</param>
        public void LoadFile(string filename)
        {
            // Clear loaded properties.
            ClearProperties();

            // Store filename.
            this.filename = filename;

            // Determine file type and load file.
            var fileInfo = new FileInfo(filename);

            // Check to make sure file actually exists.
            if (fileInfo.Exists == false)
            {
                throw new InvalidOperationException("Specified file does not exist.");
            }

            // If in multifile mode, reuse connections.
            if (multifileMode)
            {
                // Try to load file normally.  If an exception occurs, treat as a generic file.
                try
                {
                    // In multifile mode.
                    // Switch depending on file extension.
                    switch (fileInfo.Extension.ToLower())
                    {
                        case ".accdb":
                        case ".mdb":
                            // Use Dao.
                            daoFile.LoadFile(filename);
                            this.fileProperties = daoFile.FileProperties;
                            break;

                        case ".doc":
                        case ".ppt":
                        case ".xls":
                            // Use Dso.
                            dsoFile.LoadFile(filename);
                            this.fileProperties = dsoFile.FileProperties;
                            break;

                        case ".docx":
                            // Use Docx.
                            docxFile.LoadFile(filename);
                            this.fileProperties = docxFile.FileProperties;
                            break;

                        case ".pptx":
                            // Use Pptx.
                            pptxFile.LoadFile(filename);
                            this.fileProperties = pptxFile.FileProperties;
                            break;

                        case ".xlsx":
                            // Use Xlsx.
                            xlsxFile.LoadFile(filename);
                            this.fileProperties = xlsxFile.FileProperties;
                            break;

                        default:
                            // Use generic.
                            genericFile.LoadFile(filename);
                            this.fileProperties = genericFile.FileProperties;
                            break;
                    }
                }
                catch (Exception)
                {
                    // Use generic tool because of exception.
                    genericFile.LoadFile(filename, FileTypeEnum.UnknownType);
                    this.fileProperties = genericFile.FileProperties;
                }
            }
            else
            {
                // Not in multifile mode.

                // Try to load file normally.  If an exception occurs, treat as a generic file.
                try
                {
                    // Switch depending on file extension.
                    switch (fileInfo.Extension.ToLower())
                    {
                        case ".accdb":
                        case ".mdb":
                            // Use Dao.
                            this.file = new Office.Dao.DaoFile(filename);
                            break;

                        case ".doc":
                        case ".ppt":
                        case ".xls":
                            // Use Dso.
                            this.file = new Office.Dso.DsoFile(filename);
                            break;

                        case ".docx":
                            // Use Docx.
                            this.file = new Office.OpenXml.DocxFile(filename);
                            break;

                        case ".pptx":
                            // Use Pptx.
                            this.file = new Office.OpenXml.PptxFile(filename);
                            break;

                        case ".xlsx":
                            // Use Xlsx.
                            this.file = new Office.OpenXml.XlsxFile(filename);
                            break;

                        default:
                            // Use generic.
                            this.file = new Generic.GenericFile(filename);
                            break;
                    }
                }
                catch (Exception)
                {
                    // Use generic file.
                    this.file = new Generic.GenericFile(filename, FileTypeEnum.UnknownType);
                }

                // Store file properties.
                this.fileProperties = this.file.FileProperties;
            }


            // Store that file has been loaded.
            this.fileLoaded = true;
        }
    }
}
