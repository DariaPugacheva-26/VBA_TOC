Attribute VB_Name = "Table_of_contents"
Option Explicit

Dim Sh As Worksheet
Dim ShTOC As Worksheet
Dim ShName As String
Dim j As Long

Sub TOC_Tables() 'create table of contents for tables

Dim Table As ListObject
Dim TableName As String
Dim TableAddress As String 'address of first table cell for hyperlinks

j = 4 'number of start row in TOC

For Each Sh In Worksheets 'if there is alresdy such TOC then message bos and exit sub
    If Sh.Name = "Table of Contents.Tables" Then
        MsgBox "This workbook already has Table of Contents.Tables"
        Exit Sub
    End If
Next Sh

Set ShTOC = Worksheets.Add(Before:=Sheets(1))
ShTOC.Name = "Table of Contents.Tables"
ShTOC.Range("A1") = "TABLE OF CONTENTS"
ShTOC.Range("A3") = "¹"
ShTOC.Range("B3") = "Table"

For Each Sh In Worksheets
    ShName = Sh.Name
    If Sh.Name = "Table of Contents.Tables" Then GoTo nextsh
    For Each Table In Sh.ListObjects
        TableName = Table.Name
        TableAddress = Replace(Table.Range(1, 1).Address, "$", "")
        'add hyperlink
        ShTOC.Hyperlinks.Add _
        Anchor:=ShTOC.Range("B" & j), _
        Address:="", _
        SubAddress:="'" & ShName & "'!" & TableAddress, _
        TextToDisplay:=Table.Name
    
        ShTOC.Range("A" & j).Value = j - 3 'number of tables
        ShTOC.Range("A" & j).NumberFormat = "0"

        j = j + 1

    Next Table

nextsh:
Next Sh
End Sub

Sub TOC_Sheets() 'create table of contents for sheets

j = 4 'number of start row in TOC

For Each Sh In Worksheets 'if there is alresdy such TOC then message bos and exit sub
    If Sh.Name = "Table of Contents.Sheets" Then
        MsgBox "This workbook already has Table of Contents.Sheets"
    Exit Sub
    End If
Next Sh

Set ShTOC = Worksheets.Add(Before:=Sheets(1))
ShTOC.Name = "Table of Contents.Sheets"
ShTOC.Range("A1") = "TABLE OF CONTENTS"
ShTOC.Range("A3") = "¹"
ShTOC.Range("B3") = "Sheet"

For Each Sh In Worksheets
    ShName = Sh.Name
    If Sh.Name = "Table of Contents.Sheets" Then GoTo nextsh
    'add hyperlink
    ShTOC.Hyperlinks.Add _
    Anchor:=ShTOC.Range("B" & j), _
    Address:="", _
    SubAddress:="'" & ShName & "'!A1", _
    TextToDisplay:=ShName
    
    ShTOC.Range("A" & j).Value = j - 3 'number of sheets
    ShTOC.Range("A" & j).NumberFormat = "0"

    j = j + 1

nextsh:
Next Sh
End Sub


