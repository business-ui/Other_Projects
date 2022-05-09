Public Sub OrgChartRollUp()

    '
    ' This Subroutine is used to count the people in the hierarchy exported by Visio's Org Chart
    ' Two sheets in excel are returned: 
    '   Counts - gives all of the counts of each person's direct reports; sorted by Unique_ID
    '   Grouped - gives all of the counts of each person's direct reports; sorted by Reports_To; empty row between each group
    '

    Dim StartTime As Double
    StartTime = Timer
    Set xlApp = Excel.Application

    Dim workb As Workbook, works As Worksheet
    Dim tbl As ListObject
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    Dim df As Object
    Dim colNames As String

    'Converts all table objects to ranges
    For Each works In Worksheets
        For Each tbl In works.ListObjects
            tbl.Unlist
        Next tbl
    Next works

    'Checks if a previous roll-up was created
    'based on the Sheet names "Grouped" and "Counts"
    For Each works In Worksheets

        If works.Name = "Grouped" Then
            Application.DisplayAlerts = False
            works.Delete
            Application.DisplayAlerts = True
        ElseIf works.Name = "Counts" Then
            works.Name = "Sheet1"
            works.Activate
            Set ws = ActiveSheet
        End If
    Next works

    If Range("A1").Value = "Unique_ID" And Range("A2").Value = "ID1" Then
        With Range("A1")
            .Parent.ListObjects.Add(xlSrcRange, Range(.End(xlDown), .End(xlToRight)), , xlYes).Name = "Table1"
        End With
        'Setting df = "Table1"
        Set df = ws.ListObjects("Table1")

    Else
        MsgBox "Invalid Data Provided For This Function."
        Exit Sub
    End If

    'Setting up the Column
    For Each col In df.ListColumns
        colNames = colNames + CStr(col)
    Next col

    'If these Columns dont exist, add them and name them
    'If they do exist, delete them and re-add them
    If InStr(colNames, "Total") = 0 Then
        df.ListColumns.Add(5).Name = "Total"
    Else
        df.ListColumns("Total").Delete
        df.ListColumns.Add(5).Name = "Total"
    End If
    If InStr(colNames, "Contractors") = 0 Then
        df.ListColumns.Add(5).Name = "Contractors"
    Else
        df.ListColumns("Contractors").Delete
        df.ListColumns.Add(5).Name = "Contractors"
    End If
    If InStr(colNames, "FTE") = 0 Then
        df.ListColumns.Add(5).Name = "FTE"
    Else
        df.ListColumns("FTE").Delete
        df.ListColumns.Add(5).Name = "FTE"
    End If
    
    'Count the hierarchy starting with the 1st row in the DataBodyRange
    Call CountChildren(df)
    ActiveSheet.Name = "Grouped"

    'Copy the ActiveSheet
    With ActiveSheet
        .Copy After:=Sheets(Worksheets.count)
    End With
    Sheets(Worksheets.count).Name = "Counts"
    Call SortElements(df)
    Call AddSpaceBetweenGroups(df)

    'MsgBox "RunTime: " & Format((Timer - StartTime) / 86400, "hh:mm:ss") --used for recursive solution was running for an hour.

    wb.Save

End Sub

Function CountChildren(df As Object) As Collection
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
        For i  = 1 To df.ListRows.count
            Set potentialChild = df.ListColumns("Reports_To").DataBodyRange(i)

            If m = potentialChild.Value Then
                If df.ListColumns("Total").Range(potentialChild.Row) = Empty Then
                    cntChildContractor = 0
                    cntChildFTE = 0
                    cntChildTotal = 0
                Else
                    'Debug.Print df.ListColumns("Contractors").Range(potentialChild.Row)
                    cntChildContractor = df.ListColumns("Contractors").Range(potentialChild.Row).Value
                    cntChildFTE = df.ListColumns("FTE").Range(potentialChild.Row).Value
                    cntChildTotal = df.ListColumns("Total").Range(potentialChild.Row).Value
                End If

                If ((InStr(df.ListColumns("Name").Range(potentialChild.Row).Value, "(CTR)") > 0) _
                    Or (InStr(df.ListColumns("Name").Range(potentialChild.Row).Value, "- contr") > 0)) _
                    And (df.ListColumns("Title").Range(potentialChild.Row).Value <> "BOT") _
                    Then

                    df.ListColumns("Contractors").Range(currentManager).Value = df.ListColumns("Contractors").Range(currentManager).Value + cntChildContractor + 1
                    df.ListColumns("FTE").Range(currentManager).Value = df.ListColumns("FTE").Range(currentManager).Value + cntChildFTE
                    df.ListColumns("Total").Range(currentManager).Value = df.ListColumns("Total").Range(currentManager).Value + cntChildTotal + 1
                
                ElseIf df.ListColumns("Title").Range(potentialChild.Row).Value = "BOT" Then 'Skipping over any bots in the count explicitly
                Else
                    df.ListColumns("Contractors").Range(currentManager).Value = df.ListColumns("Contractors").Range(currentManager).Value + df.ListColumns("Contractors").Range(potentialChild.Row)
                    df.ListColumns("FTE").Range(currentManager).Value = df.ListColumns("FTE").Range(currentManager).Value + cntChildFTE + 1
                    df.ListColumns("Total").Range(currentManager).Value = df.ListColumns("Total").Range(currentManager).Value + cntChildTotal + 1
                End If
            End If
        Next i
    Next m

End Function

Function SortElements(df As Object)
    Dim m As Integer
    Dim i As Integer
    Dim uniqueManagers As Collection
    Set uniqueManagers = UniqueManagersCollection(df)
    Dim colNames As String

    For Each col In df.ListColumns
        colNames = colNames + CStr(col)
    Next col

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
        df.ListColumns("IdAsNum").DataBodyRange(i).Value = Mid(df.ListColumns("Unique_ID").DataBodyRange(i), 3)
    Next i

    With df.Sort
        .SortFields.Clear
        .SortFields.Add Key:=df.ListColumns("ManIdAsNum").Range, Order:=xlAscending, SortOn:=xlSortOnValues
        .SortFields.Add Key:=df.ListColumns("IdAsNum").Range, Order:=xlAscending, SortOn:=xlSortOnValues
        .Header = xlYes
        .Apply
    End With

    Dim myColl As New Collection
    Dim myArray() As Variant: ReDim myArray(0 To uniqueManagers.count)
    Dim customS As String: customS = "0,"
    Dim x As Variant
    Dim y As Variant
    myColl.Add "0"

    For Each x in df.ListColumns("IdAsNum").DataBodyRange
        For Each y in uniqueManagers
            If y = "ID" + CStr(x.Value) Then
                myColl.Add CStr(x.Value)
                customS = customS + CStr(x.Value) + ","
            End If
        Next y
    Next x

    For i = 1 To myColl.count
        'Debug.Print myColl.Item(i)
        myArray(i - 1) = myColl.Item(i)
    Next i

    customS = Left(customS, Len(customS) - 1)
    Application.AddCustomList ListArray:=myArray

    With df.Sort
        .SortFields.Clear
        '.SetRange df.DataBodyRange
        .SortFields.Add Key:=df.ListColumns("ManIdAsNum").Range, Order:=xlAscending, CustomOrder:=Application.CustomListCount, DataOption:=xlSortNormal
        '.SortFields.Add Key:=df.ListColumns("IdAsNum").Range, Order:=xlAscending, SortOn:=xlSortOnValues
        .Header = xlYes
        .Apply
    End With

    df.ListColumns("IdAsNum").Delete
    df.ListColumns("ManIdAsNum").Delete

End Function

Function AddSpaceBetweenGroups(df As Object)
    Dim i As Integer
    Dim uniqueManagers As Collection
    
    Set uniqueManagers = UniqueManagersCollection(df)

    For i = 1 To df.ListColumns("Reports_To").DataBodyRange.Count + uniqueManagers.count
        If (df.ListColumns("Reports_To").DataBodyRange(i).Value <> Empty) _
            And (df.ListColumns("Reports_To").DataBodyRange(i).Value <> df.ListColumns("Reports_To").DataBodyRange(i - 1).Value) Then
            
            df.ListRows.Add _
                Position:=i, _
                AlwaysInsert:=False
            i = i + 1
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