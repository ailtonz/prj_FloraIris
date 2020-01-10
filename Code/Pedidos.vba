Option Compare Database
Option Explicit

Private Sub cmdOBS_Padrao_Click()
Dim rObservacao As DAO.Recordset
Dim sObservacao As String

Set rObservacao = CurrentDb.OpenRecordset("Select Observacao from Observacoes")

sObservacao = rObservacao.Fields("Observacao")

Me.OBSERVACAO = sObservacao

rObservacao.Close

End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub


Private Sub cmdVisualizarPedido_Click()
    On Error GoTo Err_cmdVisualizarPedido_Click

    Dim stDocName As String

    stDocName = "Pedidos"
    
'    'Salva o registro atual
'    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    'Visualiza o pedido
    DoCmd.OpenReport stDocName, acPreview, , "Pedidos.codPedido = " & Me.codPedido

Exit_cmdVisualizarPedido_Click:
    Exit Sub

Err_cmdVisualizarPedido_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizarPedido_Click
End Sub

Private Sub codCadastro_Click()
    Me.Transportadora = Me.codCadastro.Column(2)
    Me.Transp_Telefone = Me.codCadastro.Column(3)
End Sub

Private Sub cmdImprimirPedido_Click()
On Error GoTo Err_cmdImprimirPedido_Click

    Dim stDocName As String

    stDocName = "Pedidos"
    DoCmd.OpenReport stDocName, acNormal, , "Pedidos.codPedido = " & Me.codPedido

Exit_cmdImprimirPedido_Click:
    Exit Sub

Err_cmdImprimirPedido_Click:
    MsgBox Err.Description
    Resume Exit_cmdImprimirPedido_Click
    
End Sub
Private Sub cmdNovoCadastro_Click()
On Error GoTo Err_cmdNovoCadastro_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Cadastros"
    If Not IsNull(codCadastro) Then
        stLinkCriteria = "[codCadastro] = " & codCadastro.Column(0)
        DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormEdit
    Else
        DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd
    End If
    
    AtualizarCombo = True

Exit_cmdNovoCadastro_Click:
    Exit Sub

Err_cmdNovoCadastro_Click:
    MsgBox Err.Description
    Resume Exit_cmdNovoCadastro_Click
    
End Sub

Private Sub codCadastro_GotFocus()
    codCadastro.Requery
End Sub

Private Sub Form_AfterInsert()
    Call cmdOBS_Padrao_Click
End Sub

