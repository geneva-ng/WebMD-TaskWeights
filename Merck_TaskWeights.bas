Attribute VB_Name = "Module1"
Sub Merck_TaskWeights()

Dim mycell As Range
Dim myrange As Range
Dim St As Double
Dim En As Double
Dim Offset As Double
Dim TaskNameColumn As String
Dim FinalColumn As String

'***********************************************************

'   Input these two values below to indicate the "Task Name"
'   column and final column in the spreadsheet.

    TaskNameColumn = "E"
    FinalColumn = "L"
    
'   If you run into any error messages while using this macro,
'   adjust these two values.

'***********************************************************

St = Range(TaskNameColumn & "1").Column
En = Range(FinalColumn & "1").Column
Offset = En - St + 2

Set rCell = Range(TaskNameColumn & "1")
If StrComp(rCell.Value, "Task Name") <> 0 Then
MsgBox ("The column you selected does not contain Task Names. Please refer to the Tutorial Sheet on how to fix this.")
End If

Set qCell = Range(FinalColumn & "1").Offset(0, 1)
If IsEmpty(qCell.Value) = False Then
MsgBox ("You have incorrectly indicated the final column on the sheet. Please refer to the Tutorial Sheet on how to fix this.")
Else

Set myrange = ActiveSheet.Range(TaskNameColumn & "1", Range(TaskNameColumn & "1").End(xlDown))
For Each mycell In myrange

    'header naming
    If InStr(mycell.Value, "Task Name") > 0 Then
    mycell.Offset(0, Offset).Value = "NL/WS"
    mycell.Offset(0, Offset).Font.Bold = True
      
    mycell.Offset(0, Offset + 1).Value = "CES"
    mycell.Offset(0, Offset + 1).Font.Bold = True
    
    mycell.Offset(0, Offset + 2).Value = "CT"
    mycell.Offset(0, Offset + 2).Font.Bold = True
        
    mycell.Offset(0, Offset + 3).Value = "Creative"
    mycell.Offset(0, Offset + 3).Font.Bold = True
        
    mycell.Offset(0, Offset + 4).Value = "Writer"
    mycell.Offset(0, Offset + 4).Font.Bold = True
    End If
    
    'NL/WS 1
    If InStr(mycell.Value, "Content Automation") + _
    InStr(mycell.Value, "Medical") + _
    InStr(mycell.Value, "Approval") > 0 Then
    mycell.Offset(0, Offset).Value = 1
    End If
    
    'NL/WS 2
    If InStr(mycell.Value, "Clone") + _
    InStr(mycell.Value, "Render") + _
    InStr(mycell.Value, "Newsletter Team Dev") + _
    InStr(mycell.Value, "Builds") + _
    InStr(mycell.Value, "Develops Emails") + _
    InStr(mycell.Value, "Invitations") > 0 Then
    mycell.Offset(0, Offset).Value = 2
    End If
    
    'NL/WS 3
    If InStr(mycell.Value, "Asset Review") + _
    InStr(mycell.Value, "Finalizes") + _
    InStr(mycell.Value, "Creative Formats") + _
    InStr(mycell.Value, "Abbreviated") + _
    InStr(mycell.Value, "Party HTML") + _
    InStr(mycell.Value, "Creative Format") > 0 Then
    mycell.Offset(0, Offset).Value = 3
    End If
 
    'NL/WS 4
    If InStr(mycell.Value, "Infosite Upfront Development") > 0 Then
    mycell.Offset(0, Offset).Value = 4
    End If
    
    'CES 1
    If InStr(mycell.Value, "Asset Review") + _
    InStr(mycell.Value, "Medical") + _
    InStr(mycell.Value, "Approval") + _
    InStr(mycell.Value, "Newsletter Team Dev") + _
    InStr(mycell.Value, "Finalizes") + _
    InStr(mycell.Value, "Abbreviated") + _
    InStr(mycell.Value, "Develops Driver") + _
    InStr(mycell.Value, "Develops Graphic") + _
    InStr(mycell.Value, "Develops Emails") + _
    InStr(mycell.Value, "Builds") + _
    InStr(mycell.Value, "Party HTML") + _
    InStr(mycell.Value, "Creative Format") > 0 Then
    mycell.Offset(0, Offset + 1).Value = 1
    End If
 
    'CES 2
    If InStr(mycell.Value, "Clone") > 0 Then
    mycell.Offset(0, Offset + 1).Value = 2
    End If
 
    'CES 3
    If InStr(mycell.Value, "Infosite Upfront Development") > 0 Then
    mycell.Offset(0, Offset + 1).Value = 3
    End If

    'CT 1
    If InStr(mycell.Value, "Asset Review") + _
    InStr(mycell.Value, "Approval") + _
    InStr(mycell.Value, "Party HTML") + _
    InStr(mycell.Value, "Abbreviated") + _
    InStr(mycell.Value, "Creative Format") + _
    InStr(mycell.Value, "Develops Driver") + _
    InStr(mycell.Value, "Medical") > 0 Then
    mycell.Offset(0, Offset + 2).Value = 1
    End If
    
    'Creative 2
    If InStr(mycell.Value, "Develops Graphic") + _
    InStr(mycell.Value, "Creative Clones Graphic Ads") + _
    InStr(mycell.Value, "Abbreviated") + _
    InStr(mycell.Value, "Creative Format") + _
    InStr(mycell.Value, "Party HTML") > 0 Then
    mycell.Offset(0, Offset + 3).Value = 2
    End If
 
    'Creative 4
    If InStr(mycell.Value, "Infosite Upfront Development") > 0 Then
    mycell.Offset(0, Offset + 3).Value = 4
    End If
 
    'Writer 2
    If InStr(mycell.Value, "Develops Driver") > 0 Then
    mycell.Offset(0, Offset + 4).Value = 2
    End If
 
    'Writer 4
    If InStr(mycell.Value, "Infosite Upfront Development") > 0 Then
    mycell.Offset(0, Offset + 4).Value = 4
    End If
    
Next mycell
End If
End Sub





