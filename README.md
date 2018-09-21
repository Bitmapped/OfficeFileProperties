# OfficeFileProperties
This assembly is designed to allow users to read common file properties including title, author, company, created time and modified time from common recent Office file formats.

Supported formats include:
* Microsoft Word (DOC, DOCX)
* Microsoft Excel (XLS, XLSX)
* Microsoft Access (MDB, ACCDB)
* Microsoft PowerPoint (PPT, PPTX)

For Word, Excel, and PowerPoint, this class manipulates file properties directly using DSOFile and the Open XML SDK so it is much faster than the interop classes.

## System requirements
1. NET Framework 4.5
2. One of the following installed on computer if Access databases will be used:
    * Microsoft Access MSI-based install (not Microsoft Office 365 or click-to-run-based Microsoft Office version)
    * [Microsoft Access 2016 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=54920) if Microsoft Office 365 or a click-to-run-based Microsoft Office version is not installed on the system
    * [Microsoft Access 2010 Database Engine Redistributable](https://www.microsoft.com/en-us/download/Confirmation.aspx?ID=13255) if Microsoft Office 365 or a click-to-run-based Microsoft Office version is installed on the system due to compatibility issues with click-to-run


## NuGet availability
This project is available on [NuGet](https://www.nuget.org/packages/OfficeFileProperties/).

## Usage instructions
### Getting started
1. Add **OfficeFileProperties.dll** as a reference in your project or place it in the **\bin** folder.
2. Add the dependency [**OpenXMLSDK-MOT**](https://github.com/OfficeDev/Open-XML-SDK) as a reference in your project or place its DLLs **DocumentFormat.OpenXml.dll** and **System.IO.Packaging.dll** in the **\bin** folder.
3. If you wish to use Office 97-2003 files, add the dependency [**NPOI**](https://github.com/tonyqus/npoi) as a reference in your project or place its DLL in the **\bin** folder
4. If you wish to use Access databases, ensure the Access Runtime is installed on the computer.

### Accessing file information
The below code block will show you how to access properties including creation time, modification time, author, company, title, and custom properties set on the document.
```
using (var fsFile = new OfficeFile(fullFileName))
{
    var fsProperties = fsFile.GetFileProperties();
    
    submissionFile.CreatedTimeLocal = fsProperties.CreatedTimeLocal;
    submissionFile.ModifiedTimeLocal = fsProperties.ModifiedTimeLocal;
    submissionFile.Author = fsProperties.Author;
    submissionFile.Company = fsProperties.Company;
    submissionFile.Title = fsProperties.Title;
    submissionFile.CustomProperties = fsProperties.CustomPropertiesString;
}
```

## Known issues
### BigInt type in Access 2016 v16.7 databases
The Office 365 version of Access 2016 introduced support for the [BigInt (Large Number) data type](https://support.office.com/en-us/article/Using-the-Large-Number-data-type-5b623f6e-641d-4e97-8bdf-b77bae076f70) in version 16.0.7812. When this data type is used, the database format is automatically upgraded to v16.7. MSI-based installations of Access 2016 and the Access 2016 Runtime cannot open these databases and will thrown an exception with error code `0x800A0F74`. The only solution is believed to be to install an Office 365-based version of Access on the computer.
