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

## NuGet availability
This project is available on [NuGet](https://www.nuget.org/packages/OfficeFileProperties/).

## Usage instructions
### Getting started
1. Add **OfficeFileProperties.dll** as a reference in your project.

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
