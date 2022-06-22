Public Sub SearchGALAndCount()
    Dim StartTime As Double

On Error GoTo ProcError

    Dim ol As Object
    Dim olNS As Object
    Dim olGAL As Object
    Set ol = Outlook.Session.Application
    Set olNS = ol.GetNamespace("MAPI")
    Set olGAL = olNS.AddressLists("Global Address List")
    Dim allReports As Collection
    Set allReports = New Collection

    Dim curLevelReports As Collection
    Set curLevelReports = New Collection

    Dim nextLevelReports As Collection
    Set nextLevelReports = New Collection

    Set xlApp = Excel.Application

    Dim id As Integer
    id = 1

    Dim parent_display_name As String
    parent_display_name = InputBox("Enter User Display Name")

    If StrPtr(parent_display_name) = 0 Then
        GoTo ProcCancel
    Else
        GoTo ProcContinue
    End If

ProcContinue:
    StartTime = Timer

    Dim myTopLevelReport As Outlook.ExchangeUser

    'this method returns an exchange user from their "outlook name"
    Set myTopLevelReport = olGAL.AddressEntries.Item(parent_display_name).GetExchangeUser

    Dim wb As Workbook, ws As Worksheet
    'strPath = myTopLevelReport.FirstName & " " & myTopLevelReport.LastName & ".xlsx"

    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Sheets("Sheet1")

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    For Each tbl In ws.ListObjects
        If tbl.Name = "Table1" Then ws.ListObjects("Table1").Delete
    Next tbl

    ws.Cells(1, 1).Value = "Unique_ID"
    ws.Cells(1, 2).Value = "Name"
    ws.Cells(1, 3).Value = "managerName"
    ws.Cells(1, 4).Value = "Reports_To"
    ws.Cells(1, 5).Value = "Title"
    ws.Cells(1, 6).Value = "First Name"
    ws.Cells(1, 7).Value = "Last Name"
    ws.Cells(1, 8).Value = "Office Location"
    ws.Cells(1, 9).Value = "Department"
    ws.Cells(1, 10).Value = "Email"
    'add to both the next level of reports as well as all reports
    'allReports.Add myTopLevelReport
    curLevelReports.Add myTopLevelReport

    Dim tempAddressEntries As Outlook.AddressEntries
    Dim newExUser As Outlook.ExchangeUser
    Dim i, j As Integer

    'flag for when another sublevel is found
    Dim keepLooping As Boolean
    keepLooping = False

    Dim requireValidUser As Boolean
    requireValidUser = False

    'This is where the fun begins
    Do 

        'get current reports for the current level
        For i = curLevelReports.count To 1 Step -1
            'Debug.Print curLevelReports.Item(i)
            Set tempAddressEntries = curLevelReports.Item(i).GetDirectReports

            If Range("A2").Value = "" Then
                lastRow = 1
            Else
                lastRow = Range("A1").End(xlDown).Row
            End If

            'ws.Cells(lastRow + 1, 1).Value = curLevelReports.Item(i).id
            ws.Cells(lastRow + 1, 1).Value = "ID" + CStr(GetID(id))
            ws.Cells(lastRow + 1, 2).Value = curLevelReports.Item(i).Name
            ws.Cells(lastRow + 1, 3).Value = curLevelReports.Item(i).Manager.Name
            
            ws.Cells(lastRow + 1, 5).Value = curLevelReports.Item(i).JobTitle
            ws.Cells(lastRow + 1, 6).Value = curLevelReports.Item(i).FirstName
            ws.Cells(lastRow + 1, 7).Value = curLevelReports.Item(i).LastName
            ws.Cells(lastRow + 1, 8).Value = curLevelReports.Item(i).OfficeLocation
            ws.Cells(lastRow + 1, 9).Value = curLevelReports.Item(i).Department
            ws.Cells(lastRow + 1, 10).Value = curLevelReports.Item(i).PrimarySmtpAddress

            'add all reports (note .Count returns 0 on an empty collection)
            For j = 1 To tempAddressEntries.count
                Set newExUser = tempAddressEntries.Item(j).GetExchangeUser
                nextLevelReports.Add newExUser
                keepLooping = True

                If j / 100 = 0 Then
                    wb.Save
                End If

            Next j
            Set tempAddressEntries = Nothing

        Next i

        'reset for next iteration
        Set curLevelReports = nextLevelReports
        Set nextLevelReports = New Collection

        'no more levels to keep going
        If keepLooping = False Then
            Exit Do
        End If

        'reset flag for next iteration
        keepLooping = False

    Loop

    wb.Save

    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    Call Counting
    wb.Save
    
    Exit Sub

ProcError:
    MsgBox "Invalid Outlook Name: Please Try Again."
    Exit Sub
ProcCancel:
    MsgBox "Procedure Cancelled"
    Exit Sub
End Sub

Function GetID(id As Integer) As Integer
    GetID = id
    id = id + 1
End Function

Function Counting()
    Dim StartTime As Double
    StartTime = Timer
    Set xlApp = Excel.Application

    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    Dim df As Object
    Dim colNames As String
    Dim superior As Range

    'Converting Range data to Table1 data
    If ws.Evaluate("ISREF(Table1)") Then
        Set df = ws.ListObjects("Table1")
    Else
        With Range("A1")
            .Parent.ListObjects.Add(xlSrcRange, Range(.End(xlDown), .End(xlToRight)), , xlYes).Name = "Table1"
        End With
    End If

    'Setting df = "Table1"
    Set df = ws.ListObjects("Table1")

    'Setting up the Column Header search
    For Each col In df.ListColumns
        colNames = colNames + CStr(col)
    Next col

    'If these Columns dont exist, add them and name them
    If InStr(colNames, "Total") = 0 Then
        df.ListColumns.Add(5).Name = "Total"
    End If
    If InStr(colNames, "Contractors") = 0 Then
        df.ListColumns.Add(6).Name = "Contractors"
    End If
    If InStr(colNames, "FTE") = 0 Then
        df.ListColumns.Add(7).Name = "FTE"
    End If

    For Each m In df.ListColumns("managerName").DataBodyRange
        'curr = df.ListColumns("Name").DataBodyRange.Find(What:=m, LookAt:=xlWhole).Row

        For i = 1 To df.ListRows.count
            If m.Value = df.ListColumns("Name").Range(i).Value Then
                df.ListColumns("Reports_To").Range(m.Row).Value = df.ListColumns("Unique_ID").Range(i).Value
            End If
        Next i
    Next m

    'Count the hierarchy starting with the 1st row in the DataBodyRange
    Call AltCountChildren(df)
    Call SortElements(df)
    Call AddSpaceBetweenGroups(df)
    df.DataBodyRange.Interior.ColorIndex=0
    'MsgBox "RunTime: " & Format((Timer - StartTime) / 86400, "hh:mm:ss")
    wb.Save
End Function

Function AltCountChildren(df As Object) As Collection
    'This assumes the df [DataFrame] object is sorted by Unique_ID

    Dim uniqueManagers As Collection
    Set uniqueManagers = UniqueManagersCollection(df)
    Set temp = New Collection
    Dim currentManager As Integer
    Dim potentialChild As Range
    Dim cntChildContractor As Integer
    Dim cntChildFTE As Integer
    Dim cntChildTotal As Integer

    'reverse the order of uniqueManagers
    For Each obj In uniqueManagers
        If temp.count > 0 Then
            temp.Add Item:=obj, before:=1
        Else
            temp.Add Item:=obj
        End If
    Next obj
    Set uniqueManagers = temp

    For Each m In uniqueManagers
        currentManager = df.ListColumns("Unique_ID").DataBodyRange.Find(What:=m, LookAt:=xlWhole).Row
        'Debug.Print m
        For i = 1 To df.ListRows.count
            Set potentialChild = df.ListColumns("Reports_To").DataBodyRange(i)

            If m = potentialChild.Value Then
                If df.ListColumns("Total").Range(potentialChild.Row) = Empty Then
                    cntChildContractor = 0
                    cntChildFTE = 0
                    cntChildTotal = 0
                Else
                    cntChildContractor = df.ListColumns("Contractors").Range(potentialChild.Row).Value
                    cntChildFTE = df.ListColumns("FTE").Range(potentialChild.Row).Value
                    cntChildTotal = df.ListColumns("Total").Range(potentialChild.Row).Value
                End If

                If (df.ListColumns("Title").Range(potentialChild.Row).Value <> "Contractor") _
                    And (df.ListColumns("Title").Range(potentialChild.Row).Value <> "BOT") _
                    Then
                    df.ListColumns("Contractors").Range(currentManager).Value = df.ListColumns("Contractors").Range(currentManager).Value + cntChildContractor + 1
                    df.ListColumns("FTE").Range(currentManager).Value = df.ListColumns("FTE").Range(currentManager).Value + cntChildFTE
                    df.ListColumns("Total").Range(currentManager).Value = df.ListColumns("Total").Range(currentManager).Value + cntChildTotal + 1
                ElseIf df.ListColumns("Title").Range(potentialChild.Row).Value = "BOT" Then
                Else
                    df.ListColumns("Contractors").Range(currentManager).Value = df.ListColumns("Contractors").Range(currentManager).Value + df.ListColumns("Contractors").Range(potentialChild.Row)
                    df.ListColumns("FTE").Range(currentManager).Value = df.ListColumns("FTE").Range(currentManager).Value + cntChildFTE + 1
                    df.ListColumns("Total").Range(currentManager).Value = df.ListColumns("Total").Range(currentManager).Value + cntChildTotal + 1
                End If
            End If
        Next i
    Next m

    'Dim managerFlag As Boolean
    'managerFlag = False
End Function

Function SortElements(df As Object)
    Dim m As Integer
    Dim i As Integer

    If InStr(colNames, "IdAsNum") = 0 Then
        df.ListColumns.Add(8).Name = "IdAsNum"
    End If
    If InStr(colNames, "ManIdAsNum") = 0 Then
        df.ListColumns.Add(9).Name = "ManIdAsNum"
    End If
    For m = 1 To df.ListColumns("Reports_To").DataBodyRange.count
        If df.ListColumns("Reports_To").DataBodyRange(m).Value = Empty Then
            df.ListColumns("ManIdAsNum").DataBodyRange(m).Value = 0
        Else
            df.ListColumns("ManIdAsNum").DataBodyRange(m).Value = Mid(df.ListColumns("Reports_To").DataBodyRange(m).Value, 3)
        End If
    Next m
    For i = 1 To df.ListColumns("Reports_To").DataBodyRange.count
        df.ListColumns("IdAsNum").DataBodyRange(i).Value = Mid(df.ListColumns("Unique_ID").DataBodyRange(i).Value, 3)
    Next i
    With df.Sort
        .SortFields.Clear
        .SortFields.Add Key:=df.ListColumns("ManIdAsNum").Range, Order:=xlAscending, SortOn:=xlSortOnValues
        .SortFields.Add Key:=df.ListColumns("IdAsNum").Range, Order:=xlAscending, SortOn:=xlSortOnValues
        .Header = xlYes
        .Apply
    End With
    df.ListColumns("IdAsNum").Delete
    df.ListColumns("ManIdAsNum").Delete

End Function

Function AddSpaceBetweenGroups(df As Object)
    Dim i As Integer
    For i = 1 To df.ListColumns("Reports_To").DataBodyRange.count
        i = i + 1
        If (df.ListColumns("Reports_To").DataBodyRange(i).Value <> Empty) _
            And (df.ListColumns("Reports_To").DataBodyRange(i).Value <> df.ListColumns("Reports_To").DataBodyRange(i - 1).Value) Then

            df.ListRows.Add _
                Position:=i, _
                AlwaysInsert:=False
        End If
    Next i

End Function

Function UniqueManagersCollection(df As Object) As Collection

    Dim arr As New Collection
    Dim cell As Range
    Dim duplicateFlag As Boolean
    duplicateFlag = False

    For Each cell In df.ListColumns("Reports_To").DataBodyRange
        If cell.Value = Empty Then
        Else
            For Each arrItem In arr
                If arrItem = cell.Value Then
                    duplicateFlag = True
                    Exit For
                End If
            Next arrItem
            If duplicateFlag = False Then
                arr.Add cell.Value
            End If
        End If
        duplicateFlag = False
    Next cell

    Set UniqueManagersCollection = arr

End Function