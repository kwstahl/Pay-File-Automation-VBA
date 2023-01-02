# routine-data-automation-VBA
This VBA script was created to clean routine Pay Register files. This Pay Register file was usually sent to my team from the Office of State Administration. The file was always a jumbled mess that required some advanced excel knowledge to clean.
I created This VBA script to automate the process for my team who had less advanced excel knowledge.
The data could then be used to find soldiers receiving direct deposits, soldiers receiving checks (warrants), and splits them up according to their branch of service (Army, Air, State Guards). 
This was essential, because my team handled distribution and validation of all soldier pay records/checks every pay period of 15 days.

##Cleaning Initial Data
The first button parses the data by naming the worksheet and setting all dimensions:
Setting the worksheet ensures that the data does not get mixed into wrong sheets.

`''''PREPARES ALL DATA
Dim A As Long
Dim ws As Worksheet
Dim LastCellRow As Long
Dim iCntr As Long
Set ws = ThisWorkbook.Sheets("PASTE_IN_HERE")`

The last cell is then found by using VBA's "Find" method.

`Set LastCell = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
LastCellRow = LastCell.Row`

The script checks each cell and deletes all rows where the "A" cells are empty. This is because the empty cell rows contain no useful information to be checked.

`For iCntr = LastCellRow To 1 Step -1
    If Trim(Cells(iCntr, 1)) = "" Then
        Rows(iCntr).Delete
    End If
Next`

Lastly, TextToColumns methods are performed, since each pay register file was sent as one text document. 

`Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(14, 1), Array(28, 1), Array(39, 1), Array(50, 1), _
        Array(60, 1), Array(97, 1), Array(116, 1)), TrailingMinusNumbers:=True
Sheets("PASTE_IN_HERE").Columns("A:H").AutoFit
End Sub`

##Separating Soldiers by Component, and Finding all Receiving Paper Checks (warrants)

This part names the macro, assigns dimensions, and creates new worksheets.
The sheets created are "Warrants_NG", "Warrants_SG", and "Warrants_AG". These correspond to Paychecks for Army, State, and Air guard components respectively.

`Sub GrabWarrants()
Dim activeCell1 As Range
Dim ws As Worksheet
Dim LastCellRow As Long
Set ws = ThisWorkbook.Sheets("PASTE_IN_HERE")
With ThisWorkbook.Sheets
    .Add.Name = "Warrants_NG"
    .Add.Name = "Warrants_SG"
    .Add.Name = "Warrants_AG"
End With`

The last row is then found using the same method from sub ParseRegister().

`Set LastCell = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
LastCellRow = LastCell.Row
'''' THIS IS THE LOOP TO GRAB WARRANTS`

A loop then checks each row of data and checks the "D" column for the number. 
If the number has a "1" at the beginning, then the pay type is a paycheck. 
The "H" column is checked for the service component. The first characters "SAD_TX..." will tell what service the belong to.
"SAD_TXSG*" means State Guard, and the wild card is added because the soldier name follows the component type. "SAD_TXNG" corresponds to Army, and "SAD_TXANG" corresponds to Air Force.
All records that match the criteria are then copied and pasted to their appropriate sheets.

        `For Each activeCell1 In ws.Range("D2", "D" & LastCellRow)
        WLastRow = Sheets("Warrants_SG").Range("A1").CurrentRegion.Rows.Count
            If activeCell1 Like "1########" And activeCell1.Offset(0, 4) Like "SAD_TXSG*" Then
                activeCell1.EntireRow.Copy Sheets("Warrants_SG").Range("A" & WLastRow + 1)
            End If
        Next activeCell1

        For Each activeCell1 In ws.Range("D2", "D" & LastCellRow)
        WLastRow = Sheets("Warrants_NG").Range("A1").CurrentRegion.Rows.Count
            If activeCell1 Like "1########" And activeCell1.Offset(0, 4) Like "SAD_TXNG*" Then
                activeCell1.EntireRow.Copy Sheets("Warrants_NG").Range("A" & WLastRow + 1)
            End If
        Next activeCell1
        
        For Each activeCell1 In ws.Range("D2", "D" & LastCellRow)
        WLastRow = Sheets("Warrants_AG").Range("A1").CurrentRegion.Rows.Count
            If activeCell1 Like "1########" And activeCell1.Offset(0, 4) Like "SAD_TXANG*" Then
                activeCell1.EntireRow.Copy Sheets("Warrants_AG").Range("A" & WLastRow + 1)
            End If
        Next activeCell1`
 
The final part of the code formats each sheet.       
        
`ws.Range("A1").EntireRow.Copy Sheets("Warrants_SG").Range("A1")
ws.Range("A1").EntireRow.Copy Sheets("Warrants_NG").Range("A1")
ws.Range("A1").EntireRow.Copy Sheets("Warrants_AG").Range("A1")
Sheets("Warrants_SG").Columns("A:H").AutoFit
Sheets("Warrants_NG").Columns("A:H").AutoFit
Sheets("Warrants_AG").Columns("A:H").AutoFit
End Sub`

