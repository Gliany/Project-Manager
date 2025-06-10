# Features Development Management System for Excel

This document provides the VBA code and worksheet setup required to build the "Features Development Management System" workbook. Create a new Excel **.xlsm** file and add the following sheets and VBA modules.

## 1. Data Sheet
1. Add a sheet named **Data** (hide it once the workbook is set up).
2. In `A1:B7` create a table listing the stage offsets from the Sprint Date.

```
A                B
---              ---
Brief           -90
Initiation      -80
Requirement     -60
Development       0
Testing           2
Implementation    7
```

3. Define a named range **StageOffsets** for `A1:B7`.
4. In another column (e.g., `D1:D6`) list the possible statuses: `In Process`, `Expired`, `Stuck`, `Rejected`, `Completed`, `Other`.
5. Define a named range **StatusList** for these statuses.
6. Set up named cell styles or colors for conditional formatting (e.g., red fill for expired, orange for warning, green for completed). The conditional formatting rules will use these later.
7. After setup, hide the **Data** sheet and protect it.

## 2. ProjectTemplate Sheet
1. Insert a sheet named **ProjectTemplate**. Cell `B1` should contain the Sprint Date entered by the user (`B1` is named **SprintDate** for clarity).
2. Create an Excel table named **tblStages** starting in `A3` with the following columns:
   - **StageName** – text values: Brief, Initiation, Requirement, Development, Testing, Implementation.
   - **FinalDate** – formula calculating the final date based on Sprint Date and offset.
   - **ActualDeliveryDate** – blank cells for the user or macro to fill.
   - **Status** – data validation from the **StatusList** named range.
   - **Notes** – text field.
   - **Link** – hyperlink.

Example formulas for row 2 of `tblStages` (assuming table headers are in row 3):
- `StageName`: fixed text.
- `FinalDate` (for row containing the stage name in `A4`):

```excel
=IF([@StageName]="","", SprintDate + VLOOKUP([@StageName], StageOffsets, 2, FALSE))
```

Copy this formula down the table.

### Conditional Formatting
Apply the following rules to each row of **tblStages**:
1. If the row's **Status** is not `Completed` and `TODAY()>FinalDate`, apply a red fill.
2. If the row's **Status** is not `Completed` and `FinalDate - TODAY() <= 3`, apply an orange fill.
3. If the row's **Status** is `Completed`, apply a green fill.

### Button – "Mark Stage Delivered"
Insert a form control button near the table and assign the `MarkStageDelivered` macro from Module1 (see code below).

## 3. LiveDashboard Sheet
1. Insert a sheet named **LiveDashboard**.
2. Create a table named **tblProjects** with these columns:
   - **ProjectName**
   - **SprintDate**
   - **CurrentStage**
   - **FinalDate**
   - **DeliveryDate**
   - **Status**
   - **Notes**
   - **Link**

Each row of **tblProjects** corresponds to one project sheet (every sheet except `Data`, `ProjectTemplate`, and `LiveDashboard`). The formulas below assume the project sheet's name is stored in column `A` of the table.

### Example Formulas (row 2 of tblProjects)
```
ProjectName:     =[@ProjectName]
SprintDate:      =INDIRECT("'" & [@ProjectName] & "'!B1")
CurrentStage:    =LET(stages,INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                     statuses,INDIRECT("'"&[@ProjectName]&"'!tblStages[Status]"),
                     firstIncomplete,XLOOKUP("Completed",statuses,stages,,0),
                     IF(firstIncomplete="", "Implementation", firstIncomplete))
FinalDate:       =IF([@CurrentStage]="Implementation",
                     INDIRECT("'"&[@ProjectName]&"'!tblStages[FinalDate]",TRUE),
                     XLOOKUP([@CurrentStage],
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[FinalDate]")))
DeliveryDate:    =XLOOKUP([@CurrentStage],
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[ActualDeliveryDate]"))
Status:          =XLOOKUP([@CurrentStage],
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[Status]"))
Notes:           =XLOOKUP([@CurrentStage],
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[Notes]"))
Link:            =XLOOKUP([@CurrentStage],
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[StageName]"),
                             INDIRECT("'"&[@ProjectName]&"'!tblStages[Link]"))
```
3. Apply the same conditional formatting rules to **tblProjects** as used in **tblStages** (based on the row's Status and dates).
4. Add a form control button labeled **Add New Project** and assign it to the `AddNewProject` macro from Module1.

## 4. VBA Code
Open the VBA editor (`ALT+F11`) and insert the following code.

### ThisWorkbook Module
```vb
Private Sub Workbook_Open()
    'Hide and protect the Data sheet when the workbook opens
    On Error Resume Next
    With Me.Sheets("Data")
        .Visible = xlSheetVeryHidden
        .Protect Password:="admin", UserInterfaceOnly:=True
    End With
    On Error GoTo 0
End Sub
```

### Module1 – Project Utilities
```vb
Option Explicit

'============================================================
' Adds a new project sheet based on ProjectTemplate
'============================================================
Public Sub AddNewProject()
    Dim newName As String
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    newName = InputBox("Enter new project sheet name:", "Add Project")
    If Len(newName) = 0 Then Exit Sub

    'Check for duplicate sheet name
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = newName Then
            MsgBox "A sheet with that name already exists.", vbExclamation
            Exit Sub
        End If
    Next ws

    'Copy the template
    Sheets("ProjectTemplate").Copy After:=Sheets(Sheets.Count)
    Set ws = ActiveSheet
    ws.Name = newName

    'Clear table contents except StageName and FinalDate
    With ws.ListObjects("tblStages")
        .ListColumns("ActualDeliveryDate").DataBodyRange.ClearContents
        .ListColumns("Status").DataBodyRange.ClearContents
        .ListColumns("Notes").DataBodyRange.ClearContents
        .ListColumns("Link").DataBodyRange.ClearContents
    End With

    ws.Range("B1").ClearContents 'Sprint Date

    MsgBox "Project sheet '" & newName & "' created.", vbInformation
    Sheets("LiveDashboard").Activate
    Exit Sub

ErrHandler:
    MsgBox "Error adding project: " & Err.Description, vbCritical
End Sub

'============================================================
' Marks the selected stage as delivered
'============================================================
Public Sub MarkStageDelivered()
    Dim tbl As ListObject
    Dim selRow As Long
    Dim ws As Worksheet

    On Error GoTo ErrHandler
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects("tblStages")

    If Intersect(Selection, tbl.DataBodyRange) Is Nothing Then
        MsgBox "Select a row within tblStages.", vbExclamation
        Exit Sub
    End If
    selRow = Selection.Row - tbl.DataBodyRange.Row + 1

    With tbl
        .ListColumns("ActualDeliveryDate").DataBodyRange.Cells(selRow).Value = Date
        .ListColumns("Status").DataBodyRange.Cells(selRow).Value = "Completed"
    End With
    MsgBox "Stage marked as delivered.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

## 5. Final Steps
1. Save the workbook as **macro-enabled (.xlsm)**.
2. Hide and protect the **Data** sheet (xlSheetVeryHidden) so users cannot easily unhide it.
3. Ensure all named ranges (`StageOffsets`, `StatusList`, `SprintDate`) and table names (`tblStages`, `tblProjects`) are properly defined.
4. Assign the macros to their respective buttons on the sheets.

After completing these steps, the Features Development Management System workbook will be ready for use.

