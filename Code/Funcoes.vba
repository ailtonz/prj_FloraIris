Option Compare Database
Option Explicit

Public Function ProcurarValor(codProduto As String, codPedido As Integer) As Currency

Dim Pedido As DAO.Recordset
Dim Produto As DAO.Recordset
Dim SQL_Produto As String
Dim SQL_Pedido As String

SQL_Pedido = "Select * from Pedidos where codPedido = " & codPedido
SQL_Produto = "Select * from Produtos where codProduto = " & codProduto

Set Pedido = CurrentDb.OpenRecordset(SQL_Pedido)
Set Produto = CurrentDb.OpenRecordset(SQL_Produto)

If Pedido.Fields("codTipoDePedido") = 1 Then
    ProcurarValor = Produto.Fields("Atacado")
ElseIf Pedido.Fields("codTipoDePedido") = 2 Then
    ProcurarValor = Produto.Fields("AutoAtacado")
ElseIf Pedido.Fields("codTipoDePedido") = 3 Then
    ProcurarValor = Produto.Fields("Varejo")
End If

Produto.Close
Pedido.Close

End Function

Public Function Zebrar(rpt As Report)
Static fCinza As Boolean
Const conCinza = 15198183
Const conBranco = 16777215

On Error Resume Next

    rpt.Section(0).BackColor = IIf(fCinza, conCinza, conBranco)
    fCinza = Not fCinza

End Function
