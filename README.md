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

### Included DLLs
1. DSOFile 2.1 for reading Office 97-2003 files - [Microsoft website](http://www.microsoft.com/en-us/download/details.aspx?id=8422)
2. Microsoft Office 2013 Primary Interop Assemblies for reading Access databases - [Office 2010 version](http://www.microsoft.com/en-us/download/details.aspx?id=3508)

## NuGet availability
This project is available on [NuGet](https://www.nuget.org/packages/OfficeFileProperties/).

## Usage instructions
### Getting started
1. Add **OfficeFileProperties.dll** as a reference in your project.

### Accessing file information
The below code block will show you how to access properties including creation time, modification time, author, company, title, and custom properties set on the document.
```
var fsFile = new OfficeFileProperties.File.File(fullFileName, multifileMode: false);
var CreatedDateUtc = fsFile.FileProperties.CreatedTimeUtc;
var ModifiedDateUtc = fsFile.FileProperties.ModifiedTimeUtc;

// If returned type is IOfficeFileProperties, get more properties.
if (fsFile.FileProperties is IOfficeFileProperties)
{
    var Author = ((IOfficeFileProperties)fsFile.FileProperties).Author;
    var Company = ((IOfficeFileProperties)fsFile.FileProperties).Company;
    var Title = ((IOfficeFileProperties)fsFile.FileProperties).Title;
    var CustomProperties = ((IOfficeFileProperties)fsFile.FileProperties).CustomPropertiesString;
}
```
