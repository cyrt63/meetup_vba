Attribute VB_Name = "meetupvba_1605_movie"
'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : meetupvba_1605_movie (Module)
' Project   : VBAProject
' Author    : X154045
' Date      : 15/09/2016
' Purpose   : %Description%
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1

Sub meetupvba_1605_movie_01()
      '
      '=======================================================================================
      ' Procedure : meetupvba_1605_movie_01 (Sub)
      ' Module    : meetupvba_1605_movie (Module)
      ' Project   : VBAProject
      ' Author    : X154045
      ' Date      : 15/09/2016
      ' Comments  : %Comment%
      ' Unit Test : (X154045) 15/09/2016 08:58 | Description [OK/NOK]
      ' Arg./i    :
      '           - [NO_PARAM] %Description%
      '           -
      ' Arg./o    :  ()
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '15/09/2016              X154045               Initiate
      '
      '=======================================================================================
      '

      Dim tims As Single

      Dim olist As listobject
      Dim actor_dic As New Scripting.Dictionary
      Dim NewRange As Range
      Dim orow_line() As String
      Dim orow_line2() As Variant

10        On Error GoTo Err_Handler

20    tims = Timer

30    Set olist = ActiveWorkbook.Sheets("Films_Vus").ListObjects(1)
      Dim orow As Range

40    For Each orow In olist.DataBodyRange.Rows
50        orow_line = Split(orow.Cells(, 9), ",")
          
60        For j = LBound(orow_line) To UBound(orow_line)
70            If orow_line(j) <> "" Then
80                Select Case actor_dic.Exists(CStr(Trim(orow_line(j))))
                  Case False
90                        ReDim orow_line2(1, 3)
100                       orow_line2(1, 1) = Trim(orow_line(j))
110                       orow_line2(1, 2) = orow.Cells(, 1)
120                       orow_line2(1, 3) = 1
130                       actor_dic.Add CStr(Trim(orow_line(j))), orow_line2
140               Case True
150                       ReDim orow_line2(1, 3)
160                       orow_line2 = actor_dic.Item(Trim(orow_line(j)))
170                       orow_line2(1, 2) = orow_line2(1, 2) & "," & orow.Cells(, 1)
180                       orow_line2(1, 3) = orow_line2(1, 3) + 1
190                       actor_dic.Item(Trim(orow_line(j))) = orow_line2
                          
200               End Select
210           End If
220       Next
230   Next

240   Set olist = ActiveWorkbook.Sheets("acteur").ListObjects(1)
250   If olist.ListRows.Count > 0 Then olist.DataBodyRange.Delete
260   With olist.Parent
          Dim orow_line_data
270       orow_line_data = Application.Transpose(Application.Transpose(actor_dic.Items))
280       Set NewRange = .Range(Cells(2, 1), Cells(UBound(orow_line_data), UBound(orow_line_data, 2)))
290       NewRange = orow_line_data
300   End With

310   Debug.Print "time lpas ", Timer - tims

Err_Exit:
320       Set olist = Nothing
330       Exit Sub
          
Err_Handler:
340       GoTo Err_Exit

End Sub


Sub meetupvba_1605_movie_02()
      '
      '=======================================================================================
      ' Procedure : meetupvba_1605_movie_02 (Sub)
      ' Module    : meetupvba_1605_movie (Module)
      ' Project   : VBAProject
      ' Author    : X154045
      ' Date      : 15/09/2016
      ' Comments  : %Comment%
      ' Unit Test : (X154045) 15/09/2016 09:00 | Description [OK]
      ' Arg./i    :
      '           - [%PARAM1%] %Description%
      '           -
      ' Arg./o    :  ()
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '15/09/2016              X154045               Initiate
      '
      '=======================================================================================
      '
      Dim tims As Single
      Dim olist As listobject
      Dim orng As Range
      Dim actors_listr() As Variant
      Dim actors_list() As String
      Dim osplit() As String
      Dim movie_list() As Variant
      Dim actors_dic As New Scripting.Dictionary
      Dim count_dic As New Scripting.Dictionary

10        On Error GoTo Err_Handler

20    tims = Timer

30    Set olist = ActiveWorkbook.Sheets("Films_Vus").ListObjects(1)

40    Set orng = olist.DataBodyRange.Columns(9)
50    actors_listr = Application.Transpose(orng)
60    actors_list = Filter(Split(Join(Application.Index(actors_listr, 1, 0), ","), ","), " ", True)

70    Set orng = olist.DataBodyRange.Columns(1)
80    movie_list = Application.Transpose(orng)
90    Set olist = Nothing

100   For i = LBound(actors_listr) To UBound(actors_listr)
110       osplit = Filter(Split(actors_listr(i), ","), " ", True)
120       For j = LBound(osplit) To UBound(osplit)
          
130           actor = CStr(Trim(osplit(j)))
              
140           If actors_dic.Exists(actor) = False Then
150               actors_dic.Add actor, movie_list(i)
160               count_dic.Add actor, 1
170           Else
180               actors_dic.Item(actor) = actors_dic.Item(actor) & "," & movie_list(i)
190               count_dic.Item(actor) = count_dic.Item(actor) + 1
200           End If
              
210       Next j
              
220   Next i

230   Set olist = ActiveWorkbook.Sheets("acteur").ListObjects(1)
240   If olist.ListRows.Count > 0 Then olist.DataBodyRange.Delete
250   With olist.Parent
          Dim orow_line_data
260       Set orng = .Range(Cells(2, 1), Cells(actors_dic.Count, 1))
270       orng = Application.Transpose(actors_dic.Keys)
280       Set orng = .Range(Cells(2, 2), Cells(actors_dic.Count, 2))
290       orng = Application.Transpose(actors_dic.Items)
300       Set orng = .Range(Cells(2, 3), Cells(actors_dic.Count, 3))
310       orng = Application.Transpose(count_dic.Items)
          
320   End With

330   Debug.Print "time laps ", Timer - tims

Err_Exit:
340       Set olist = Nothing
350       Set orng = Nothing
360       Exit Sub
          
Err_Handler:
370       GoTo Err_Exit

End Sub
