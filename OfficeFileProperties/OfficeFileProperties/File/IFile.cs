using System;
namespace OfficeFileProperties.File
{
    interface IFile
    {
        IFileProperties FileProperties { get; }
        void LoadFile(string filename);
    }
}
