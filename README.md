# C# Report Generator

A Windows Forms application for generating HTML reports from SQL Server stored procedures with email distribution and Excel export.

## Overview

This is sanitized sample code from a reporting application I built at a financial services firm. It demonstrates:

- **Threading** - Background wait dialogs using STA threads
- **Reflection** - Dynamic method invocation for embedded system functions
- **DataTable manipulation** - Filtering, sorting, grouping with DataView
- **HTML templating** - Report generation with summary/detail sections
- **Email with attachments** - SMTP with Excel/text attachments via MemoryStream
- **Active Directory integration** - Recursive group membership enumeration
- **XML configuration** - LINQ to XML for parsing connection/report metadata

## Files

### Reporter.cs
Core report generation engine:
- `GenerateReport()` - Main entry point, orchestrates the full pipeline
- Threading with `STA` apartment state for Windows Forms compatibility
- DataTable grouping with `Compute()` for aggregate counts
- Dynamic function execution using `MethodInfo.Invoke()`

### DataAccess.cs
Data access and output:
- `ExecSQL()` - Query execution with table-valued parameter support
- `ExcelFile()` - DataSet to Excel using EPPlus library
- `SendEmail()` - HTML email with MemoryStream attachments
- XML parsing for connection discovery

### SystemLibrary.cs
Utility functions (called via reflection):
- `CurrentDatetime()` - Formatted timestamp
- `FirstOfMonth()` - Date calculation
- `ADGroupMembers()` - Recursive AD group enumeration

## Key Patterns Demonstrated

### Reflection for Dynamic Function Calls
```csharp
Type type = typeof(SystemLibrary);
MethodInfo methodInfo = type.GetMethod(function);
returnValue = (string)methodInfo.Invoke(null, parametersArray);
```

### Threading with Windows Forms
```csharp
Thread t = new Thread(Wait);
t.SetApartmentState(ApartmentState.STA);
t.Start();
```

### DataTable Grouping
```csharp
DataView dvGroup = new DataView(detail);
DataTable dtGroup = dvGroup.ToTable(true, groupBy);
dr["Count"] = detail.Compute("Count(" + countColumn + ")", filter);
```

### Email with MemoryStream Attachment
```csharp
using (MemoryStream memoryStream = new MemoryStream())
{
    ExcelPackage contentAsExcel = ExcelFile(AttachmentDataSet);
    contentAsExcel.SaveAs(memoryStream);
    memoryStream.Seek(0, SeekOrigin.Begin);
    Attachment attachment = new Attachment(memoryStream, contentType);
}
```

## Technology Stack

- .NET Framework 4.x
- Windows Forms
- System.Data.SqlClient
- System.DirectoryServices
- System.Net.Mail
- EPPlus (Excel generation)
- LINQ to XML

## Architecture

```
[Command Line] -> [Reporter.cs] -> [DataAccess.cs] -> [SQL Server]
                      |                  |
                      v                  v
               [HTML Output]      [Email/Excel]
                      |
                      v
               [ReportViewer Form]
```

## Author

Stephen Lantz - Senior Database Engineer
20+ years SQL Server, PostgreSQL, ETL, and data architecture

## Note

This is representative sample code, sanitized from proprietary systems. It demonstrates coding patterns and style rather than a complete runnable application.
