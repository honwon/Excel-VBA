VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 유저폼 
   Caption         =   "드래곤이 되자!"
   ClientHeight    =   5640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5160
   OleObjectBlob   =   "유저폼.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "유저폼"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
Dim Rng As Range, hh As Range

If Range("C1") = "" Then
  MsgBox "정렬할 데이터 없음!"
Else

  n = Range("C1").End(xlToRight).Column
  Set Rng = Range("B:B")
  
  For i = 3 To n
    If Cells(1, i) = "-" Then
      a = Cells(1, i).End(4).Row
      Set hh = Range(Cells(1, i), Cells(a, i))
      Set Rng = Union(Rng, hh)
    End If
  Next i
  Rng.Select
  Selection.EntireColumn.Delete

End If
End Sub

Private Sub cmd2_Click()
Dim a As Range

If Range("C2") = "" Then
  MsgBox "복사할 데이터 없음!"
Else
    Sheets("데이터 정렬 (C1에 복사)").Select
  Range("C2").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Set a = Range(Selection, Selection.End(xlDown))
  Set b = a.Resize(a.Rows.Count, a.Columns.Count - 1)
  b.Select
  Selection.Copy
  Sheets("전체 데이터").Select
  Range("전체_데이터[[#Headers],[주차]]").Select
  Selection.End(xlDown).Offset(1, 1).Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  
End If

End Sub

Private Sub cmd3_Click()

If Range("C2") = "" Then
  MsgBox "복사할 데이터 없음!"
Else
    Sheets("데이터 정렬 (C1에 복사)").Select
    Range("C2").Select
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("다이어트 기록").Select
    Range("표1_4[[#Headers],[성공여부]]").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End If

End Sub

Private Sub cmd4_Click()
    Sheets("데이터 정렬 (C1에 복사)").Select
End Sub

Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub CommandButton2_Click()
If Range("C2") = "" Then
  MsgBox "삭제할 데이터 없음!"
Else
    Range("C2").CurrentRegion.Select
    Selection.ClearContents
    Range("B1").Select
End If

End Sub
