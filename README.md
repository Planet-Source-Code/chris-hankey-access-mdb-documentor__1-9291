<div align="center">

## Access/MDB documentor


</div>

### Description

Excel macro that extracts all tables, fields, field types, queries & descriptions from a JET/Access database.

Very useful for documenting Access databases.
 
### More Info
 
msgbox will prompt for a path to the database.

This macro was developed with DAO 3.51, but should work with any of the later versions of DAO. It has not been tested with Access 2000.

Paste the code into a module and make sure to set a reference to the DAO library.

This code populates a spreadsheet with schema info.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[chris hankey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-hankey.md)
**Level**          |Beginner
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-hankey-access-mdb-documentor__1-9291/archive/master.zip)





### Source Code

```
Sub GetMDBDescription()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Creator  chris hankey
'Inputs   none
'Returns  none
'Created  1/14/2000
'Modified
'Notes   extracts all field and table descriptions from the database
'      indicated by the user and loads them into the active sheet.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim sPath As String
  Dim db As Database
  Dim tdf As TableDef
  Dim qdf As QueryDef
  Dim fld As Field
  Dim iRow As Integer
  Dim sTemp As String
  On Error GoTo ErrorHandler
  'get the path of the mdb from the user
  sPath = InputBox("Please enter the MDB's path")
  'clear the sheets contents. Also removes all formatting
  Cells.Delete
  iRow = 1
  'exit the sub if the user does not enter a path
  If sPath <> vbNullString Then
    'test the path to make sure that it actually points to a file
    sPathTest = Dir(sPath, vbNormal)
    Set db = OpenDatabase(sPath)
    'format the sheet now that we have received a valid MDB
    'to open
    Columns("A:A").VerticalAlignment = xlTop
    Columns("A:A").ColumnWidth = 36
    Columns("B:B").VerticalAlignment = xlTop
    Columns("B:B").WrapText = True
    Columns("B:B").ColumnWidth = 26
    Columns("D:D").VerticalAlignment = xlTop
    Columns("D:D").WrapText = True
    Columns("D:D").ColumnWidth = 43
    ActiveSheet.Cells(iRow, 1) = "Tables"
    ActiveSheet.Cells(iRow, 1).Font.Bold = True
    ActiveSheet.Cells(iRow, 1).Font.Size = 12
    iRow = iRow + 1
    'scroll thru the tabledefs
    For Each tdf In db.TableDefs
      'skip Access System tables - they all start with MSys
      If Left(tdf.Name, 4) <> "MSys" Then
        ActiveSheet.Cells(iRow, 1) = tdf.Name
        ActiveSheet.Cells(iRow, 1).Font.Bold = True
        ActiveSheet.Cells(iRow, 1).Font.Underline = xlUnderlineStyleSingle
        ActiveSheet.Cells(iRow, 2) = tdf.Properties("Description")
        'merge the cells for the table descriptions
        sTemp = "B" & iRow & ":D" & iRow
        Range(sTemp).MergeCells = True
        iRow = iRow + 1
        'generate a header for the fields
        ActiveSheet.Cells(iRow, 2) = "Field Name"
        ActiveSheet.Cells(iRow, 2).Font.Bold = True
        ActiveSheet.Cells(iRow, 2).Font.Underline = xlUnderlineStyleSingle
        ActiveSheet.Cells(iRow, 3) = "Type"
        ActiveSheet.Cells(iRow, 3).Font.Bold = True
        ActiveSheet.Cells(iRow, 3).Font.Underline = xlUnderlineStyleSingle
        ActiveSheet.Cells(iRow, 4) = "Description"
        ActiveSheet.Cells(iRow, 2).Font.Bold = True
        ActiveSheet.Cells(iRow, 4).Font.Underline = xlUnderlineStyleSingle
        iRow = iRow + 1
        'scroll thru the fields
        For Each fld In tdf.Fields
          ActiveSheet.Cells(iRow, 2) = fld.Name
          ActiveSheet.Cells(iRow, 2).Font.Bold = True
          ActiveSheet.Cells(iRow, 3) = TypeName(fld.Type)
          ActiveSheet.Cells(iRow, 4) = fld.Properties("Description")
          iRow = iRow + 1
        Next fld
        iRow = iRow + 1
      End If
    Next tdf
    'generate a query section header
    iRow = iRow + 1
    ActiveSheet.Cells(iRow, 1) = "Queries"
    ActiveSheet.Cells(iRow, 1).Font.Bold = True
    ActiveSheet.Cells(iRow, 1).Font.Size = 12
    'merge the cells for the Query descriptions
    sTemp = "B" & iRow & ":D" & iRow
    Range(sTemp).MergeCells = True
    iRow = iRow + 1
    'scroll thru the queries
    For Each qdf In db.QueryDefs
      ActiveSheet.Cells(iRow, 1) = qdf.Name
      ActiveSheet.Cells(iRow, 1).Font.Bold = True
      ActiveSheet.Cells(iRow, 1).Font.Underline = xlUnderlineStyleSingle
      ActiveSheet.Cells(iRow, 4) = qdf.Properties("Description")
      'merge the cells for the Query descriptions
      sTemp = "B" & iRow & ":D" & iRow
      Range(sTemp).MergeCells = True
      iRow = iRow + 1
    Next qdf
  End If
ExitSub:
  Exit Sub
ErrorHandler:
  Select Case Err
    Case 3270 'property not found
      Resume Next
    Case Else
      MsgBox Err.Description
      GoTo ExitSub
  End Select
End Sub
Function TypeName(iType As Integer) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Creator  chris hankey
'Inputs   iType - data type of field
'Returns  string containing name of type
'Created  1/14/2000
'Modified
'Notes
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Select Case iType
    Case dbBigInt
      TypeName = "Big Integer"
    Case dbBinary
      TypeName = "Binary"
    Case dbBoolean
      TypeName = "Boolean"
    Case dbByte
      TypeName = "Byte"
    Case dbChar
      TypeName = "Char"
    Case dbCurrency
      TypeName = "Currency"
    Case dbDate
      TypeName = "Date"
    Case dbDecimal
      TypeName = "Decimal"
    Case dbDouble
      TypeName = "Double"
    Case dbFloat
      TypeName = "Float"
    Case dbGUID
      TypeName = "GUID"
    Case dbInteger
      TypeName = "Integer"
    Case dbLong
      TypeName = "Long"
    Case dbLongBinary
      TypeName = "Long Binary"
    Case dbMemo
      TypeName = "Memo"
    Case dbNumeric
      TypeName = "Numeric"
    Case dbSingle
      TypeName = "Single"
    Case dbText
      TypeName = "Text"
    Case dbTime
      TypeName = "Time"
    Case dbTimeStamp
      TypeName = "Time Stamp"
    Case dbVarBinary
      TypeName = "VarBinary"
    Case Else
      TypeName = ""
  End Select
End Function
```

