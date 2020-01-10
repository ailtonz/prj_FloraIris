Option Compare Database
Option Explicit
Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Fechamento"
    
    DoCmd.OpenReport stDocName, acPreview, , stLinkCriteria

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    Me.Inicio = Format(Now(), "dd/mm/yyyy")
    Me.Terminio = Format(Now(), "dd/mm/yyyy")
    
End Sub
