# OfficeFileProperties
This assembly is designed to allow users to read common file properties including title, author, company, created time and modified time from common recent Office file formats.

Supported formats include:
* Microsoft Word (DOC, DOCX)
* Microsoft Excel (XLS, XLSX)
* Microsoft Access (MDB, ACCDB)
* Microsoft PowerPoint (PPT, PPTX)

For Word, Excel, and PowerPoint, this class manipulates file properties directly using DSOFile and the Open XML SDK so it is much faster than the interop classes.

## System Requirements
1. Microsoft .NET Framework 4
2. DSOFile 2.1 for reading Office 97-2003 files - http://www.microsoft.com/en-us/download/details.aspx?id=8422
3. Microsoft Office 2010 Primary Interop Assemblies for reading Access databases - http://www.microsoft.com/en-us/download/details.aspx?id=3508
