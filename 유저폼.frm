VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ������ 
   Caption         =   "�巡���� ����!"
   ClientHeight    =   5640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5160
   OleObjectBlob   =   "������.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
Dim Rng As Range, hh As Range

If Range("C1") = "" Then
  MsgBox "������ ������ ����!"
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
  MsgBox "������ ������ ����!"
Else
    Sheets("������ ���� (C1�� ����)").Select
  Range("C2").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Set a = Range(Selection, Selection.End(xlDown))
  Set b = a.Resize(a.Rows.Count, a.Columns.Count - 1)
  b.Select
  Selection.Copy
  Sheets("��ü ������").Select
  Range("��ü_������[[#Headers],[����]]").Select
  Selection.End(xlDown).Offset(1, 1).Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  
End If

End Sub

Private Sub cmd3_Click()

If Range("C2") = "" Then
  MsgBox "������ ������ ����!"
Else
    Sheets("������ ���� (C1�� ����)").Select
    Range("C2").Select
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("���̾�Ʈ ���").Select
    Range("ǥ1_4[[#Headers],[��������]]").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End If

End Sub

Private Sub cmd4_Click()
    Sheets("������ ���� (C1�� ����)").Select
End Sub

Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub CommandButton2_Click()
If Range("C2") = "" Then
  MsgBox "������ ������ ����!"
Else
    Range("C2").CurrentRegion.Select
    Selection.ClearContents
    Range("B1").Select
End If

End Sub
