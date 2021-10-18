VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ws2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()

MsgBox "TIC TAC TOE - let's play! You need 5 (X or Y) in a row to win"

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

ws3.Range("A1") = "Y"

Dim LastRow As Integer
LastRow = ws3.Cells(Rows.Count, 1).End(xlUp).Row

On Error GoTo Handler

    If Selection.Cells.Count > 1 Then
        MsgBox "You can select only 1 cell"
    ElseIf Target <> vbNullString Then Exit Sub
    Else
        If ws3.Range("A" & LastRow) = "X" Then
            Target = "Y"
            ws3.Range("A" & LastRow + 1) = "Y"
                

                If (Target.Offset(1).Value = "Y" And Target.Offset(2).Value = "Y" And Target.Offset(3).Value = "Y" And Target.Offset(4).Value = "Y") _
                    Or (Target.Offset(-1).Value = "Y" And Target.Offset(1).Value = "Y" And Target.Offset(2).Value = "Y" And Target.Offset(3).Value = "Y") _
                    Or (Target.Offset(-2).Value = "Y" And Target.Offset(-1).Value = "Y" And Target.Offset(1).Value = "Y" And Target.Offset(2).Value = "Y") _
                    Or (Target.Offset(-3).Value = "Y" And Target.Offset(-2).Value = "Y" And Target.Offset(-1).Value = "Y" And Target.Offset(1).Value = "Y") _
                    Or (Target.Offset(-1).Value = "Y" And Target.Offset(-2).Value = "Y" And Target.Offset(-3).Value = "Y" And Target.Offset(-4).Value = "Y") Then
                    MsgBox "Y won!"
                    Call Clearing
                    Exit Sub
                End If
                

                If (Target.Offset(0, 1).Value = "Y" And Target.Offset(0, 2).Value = "Y" And Target.Offset(0, 3).Value = "Y" And Target.Offset(0, 4).Value = "Y") _
                    Or (Target.Offset(0, -1).Value = "Y" And Target.Offset(0, 1).Value = "Y" And Target.Offset(0, 2).Value = "Y" And Target.Offset(0, 3).Value = "Y") _
                    Or (Target.Offset(0, -2).Value = "Y" And Target.Offset(0, -1).Value = "Y" And Target.Offset(0, 1).Value = "Y" And Target.Offset(0, 2).Value = "Y") _
                    Or (Target.Offset(0, -3).Value = "Y" And Target.Offset(0, -2).Value = "Y" And Target.Offset(0, -1).Value = "Y" And Target.Offset(0, 1).Value = "Y") _
                    Or (Target.Offset(0, -1).Value = "Y" And Target.Offset(0, -2).Value = "Y" And Target.Offset(0, -3).Value = "Y" And Target.Offset(0, -4).Value = "Y") Then
                    MsgBox "Y won!"
                    Call Clearing
                    Exit Sub
                End If

                If (Target.Offset(1, 1).Value = "Y" And Target.Offset(2, 2).Value = "Y" And Target.Offset(3, 3).Value = "Y" And Target.Offset(4, 4).Value = "Y") _
                    Or (Target.Offset(-1, -1).Value = "Y" And Target.Offset(1, 1).Value = "Y" And Target.Offset(2, 2).Value = "Y" And Target.Offset(3, 3).Value = "Y") _
                    Or (Target.Offset(-2, -2).Value = "Y" And Target.Offset(-1, -1).Value = "Y" And Target.Offset(1, 1).Value = "Y" And Target.Offset(2, 2).Value = "Y") _
                    Or (Target.Offset(-3, -3).Value = "Y" And Target.Offset(-2, -2).Value = "Y" And Target.Offset(-1, -1).Value = "Y" And Target.Offset(1, 1).Value = "Y") _
                    Or (Target.Offset(-4, -4).Value = "Y" And Target.Offset(-3, -3).Value = "Y" And Target.Offset(-2, -2).Value = "Y" And Target.Offset(-1, -1).Value = "Y") _
                    Or (Target.Offset(1, 1).Value = "Y" And Target.Offset(-1, -1).Value = "Y" And Target.Offset(-2, -2).Value = "Y" And Target.Offset(-3, -3).Value = "Y") _
                    Or (Target.Offset(2, 2).Value = "Y" And Target.Offset(1, 1).Value = "Y" And Target.Offset(-1, -1).Value = "Y" And Target.Offset(-2, -2).Value = "Y") _
                    Or (Target.Offset(3, 3).Value = "Y" And Target.Offset(2, 2).Value = "Y" And Target.Offset(1, 1).Value = "Y" And Target.Offset(-1, -1).Value = "Y") _
                    Or (Target.Offset(-1, 1).Value = "Y" And Target.Offset(-2, 2).Value = "Y" And Target.Offset(-3, 3).Value = "Y" And Target.Offset(-4, 4).Value = "Y") _
                    Or (Target.Offset(1, -1).Value = "Y" And Target.Offset(2, -2).Value = "Y" And Target.Offset(3, -3).Value = "Y" And Target.Offset(4, -4).Value = "Y") Then
                    MsgBox "Y won!"
                    Call Clearing
                    Exit Sub
                End If

        Else
            Target = "X"
            ws3.Range("A" & LastRow + 1) = "X"
            
                If (Target.Offset(1).Value = "X" And Target.Offset(2).Value = "X" And Target.Offset(3).Value = "X" And Target.Offset(4).Value = "X") _
                    Or (Target.Offset(-1).Value = "X" And Target.Offset(1).Value = "X" And Target.Offset(2).Value = "X" And Target.Offset(3).Value = "X") _
                    Or (Target.Offset(-2).Value = "X" And Target.Offset(-1).Value = "X" And Target.Offset(1).Value = "X" And Target.Offset(2).Value = "X") _
                    Or (Target.Offset(-3).Value = "X" And Target.Offset(-2).Value = "X" And Target.Offset(-1).Value = "X" And Target.Offset(1).Value = "X") _
                    Or (Target.Offset(-1).Value = "X" And Target.Offset(-2).Value = "X" And Target.Offset(-3).Value = "X" And Target.Offset(-4).Value = "X") Then
                    MsgBox "X won!"
                    Call Clearing
                    Exit Sub
                End If
                


                If (Target.Offset(0, 1).Value = "X" And Target.Offset(0, 2).Value = "X" And Target.Offset(0, 3).Value = "X" And Target.Offset(0, 4).Value = "X") _
                    Or (Target.Offset(0, -1).Value = "X" And Target.Offset(0, 1).Value = "X" And Target.Offset(0, 2).Value = "X" And Target.Offset(0, 3).Value = "X") _
                    Or (Target.Offset(0, -2).Value = "X" And Target.Offset(0, -1).Value = "X" And Target.Offset(0, 1).Value = "X" And Target.Offset(0, 2).Value = "X") _
                    Or (Target.Offset(0, -3).Value = "X" And Target.Offset(0, -2).Value = "X" And Target.Offset(0, -1).Value = "X" And Target.Offset(0, 1).Value = "X") _
                    Or (Target.Offset(0, -1).Value = "X" And Target.Offset(0, -2).Value = "X" And Target.Offset(0, -3).Value = "X" And Target.Offset(0, -4).Value = "X") Then
                    MsgBox "X won!"
                    Call Clearing
                    Exit Sub
                End If
                


                If (Target.Offset(1, 1).Value = "X" And Target.Offset(2, 2).Value = "X" And Target.Offset(3, 3).Value = "X" And Target.Offset(4, 4).Value = "X") _
                    Or (Target.Offset(-1, -1).Value = "X" And Target.Offset(1, 1).Value = "X" And Target.Offset(2, 2).Value = "X" And Target.Offset(3, 3).Value = "X") _
                    Or (Target.Offset(-2, -2).Value = "X" And Target.Offset(-1, -1).Value = "X" And Target.Offset(1, 1).Value = "X" And Target.Offset(2, 2).Value = "X") _
                    Or (Target.Offset(-3, -3).Value = "X" And Target.Offset(-2, -2).Value = "X" And Target.Offset(-1, -1).Value = "X" And Target.Offset(1, 1).Value = "X") _
                    Or (Target.Offset(-4, -4).Value = "X" And Target.Offset(-3, -3).Value = "X" And Target.Offset(-2, -2).Value = "X" And Target.Offset(-1, -1).Value = "X") _
                    Or (Target.Offset(1, 1).Value = "X" And Target.Offset(-1, -1).Value = "X" And Target.Offset(-2, -2).Value = "X" And Target.Offset(-3, -3).Value = "X") _
                    Or (Target.Offset(2, 2).Value = "X" And Target.Offset(1, 1).Value = "X" And Target.Offset(-1, -1).Value = "X" And Target.Offset(-2, -2).Value = "X") _
                    Or (Target.Offset(3, 3).Value = "X" And Target.Offset(2, 2).Value = "X" And Target.Offset(1, 1).Value = "X" And Target.Offset(-1, -1).Value = "X") _
                    Or (Target.Offset(-1, 1).Value = "X" And Target.Offset(-2, 2).Value = "X" And Target.Offset(-3, 3).Value = "X" And Target.Offset(-4, 4).Value = "X") _
                    Or (Target.Offset(1, -1).Value = "X" And Target.Offset(2, -2).Value = "X" And Target.Offset(3, -3).Value = "X" And Target.Offset(4, -4).Value = "X") Then
                    MsgBox "X won!"
                    Call Clearing
                    Exit Sub
                End If
                           
        End If
    End If

Handler:
    
End Sub

Sub Restart_Game()

Call Clearing

MsgBox "TIC TAC TOE - let's play! You need 5 (X or Y) in a row to win"

End Sub

Sub Clearing()

        ws2.Range("F6:AS35").ClearContents
        ws3.Cells.ClearContents
                    
End Sub
