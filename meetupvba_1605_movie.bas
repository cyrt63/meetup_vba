Attribute VB_Name = "mMovie"
Option Base 1


Sub test_1()
Dim tims As Single

Dim olist As listobject
Dim actor_dic As New Scripting.Dictionary
Dim NewRange As Range
Dim orow_line() As String
Dim orow_line2() As Variant

tims = Timer

Set olist = ActiveWorkbook.Sheets("Films_Vus").ListObjects(1)
Dim orow As Range

For Each orow In olist.DataBodyRange.Rows
    orow_line = Split(orow.Cells(, 9), ",")
    
    For j = LBound(orow_line) To UBound(orow_line)
        If orow_line(j) <> "" Then
            Select Case actor_dic.Exists(CStr(Trim(orow_line(j))))
            Case False
                    ReDim orow_line2(1, 3)
                    orow_line2(1, 1) = Trim(orow_line(j))
                    orow_line2(1, 2) = orow.Cells(, 1)
                    orow_line2(1, 3) = 1
                    actor_dic.Add CStr(Trim(orow_line(j))), orow_line2
            Case True
                    ReDim orow_line2(1, 3)
                    orow_line2 = actor_dic.Item(Trim(orow_line(j)))
                    orow_line2(1, 2) = orow_line2(1, 2) & "," & orow.Cells(, 1)
                    orow_line2(1, 3) = orow_line2(1, 3) + 1
                    actor_dic.Item(Trim(orow_line(j))) = orow_line2
                    
            End Select
        End If
    Next
Next

Set olist = ActiveWorkbook.Sheets("acteur").ListObjects(1)
If olist.ListRows.Count > 0 Then olist.DataBodyRange.Delete
With olist.Parent
    Dim orow_line_data
    orow_line_data = Application.Transpose(Application.Transpose(actor_dic.Items))
    Set NewRange = .Range(Cells(2, 1), Cells(UBound(orow_line_data), UBound(orow_line_data, 2)))
    NewRange = orow_line_data
End With

Debug.Print "time lpas ", Timer - tims
End Sub


Sub test_2()
Dim tims As Single
Dim olist As listobject
Dim orng As Range
Dim actors_listr() As Variant
Dim actors_list() As String
Dim osplit() As String
Dim movie_list() As Variant
Dim actors_dic As New Scripting.Dictionary
Dim count_dic As New Scripting.Dictionary

tims = Timer

Set olist = ActiveWorkbook.Sheets("Films_Vus").ListObjects(1)

Set orng = olist.DataBodyRange.Columns(9)
actors_listr = Application.Transpose(orng)
actors_list = Filter(Split(Join(Application.Index(actors_listr, 1, 0), ","), ","), " ", True)

Set orng = olist.DataBodyRange.Columns(1)
movie_list = Application.Transpose(orng)
Set olist = Nothing

For i = LBound(actors_listr) To UBound(actors_listr)
    osplit = Filter(Split(actors_listr(i), ","), " ", True)
    For j = LBound(osplit) To UBound(osplit)
    
        actor = CStr(Trim(osplit(j)))
        
        If actors_dic.Exists(actor) = False Then
            actors_dic.Add actor, movie_list(i)
            count_dic.Add actor, 1
        Else
            actors_dic.Item(actor) = actors_dic.Item(actor) & "," & movie_list(i)
            count_dic.Item(actor) = count_dic.Item(actor) + 1
        End If
        
    Next j
        
Next i

Set olist = ActiveWorkbook.Sheets("acteur").ListObjects(1)
If olist.ListRows.Count > 0 Then olist.DataBodyRange.Delete
With olist.Parent
    Dim orow_line_data
    Set orng = .Range(Cells(2, 1), Cells(actors_dic.Count, 1))
    orng = Application.Transpose(actors_dic.Keys)
    Set orng = .Range(Cells(2, 2), Cells(actors_dic.Count, 2))
    orng = Application.Transpose(actors_dic.Items)
    Set orng = .Range(Cells(2, 3), Cells(actors_dic.Count, 3))
    orng = Application.Transpose(count_dic.Items)
    
End With

Debug.Print "time laps ", Timer - tims
End Sub
