Option Compare Database
Option Explicit

Private Sub Quantidade_AfterUpdate()
    
    If Me.Quantidade > 0 Then
        Me.ValorUnitario = ProcurarValor(Me.codProduto, Me.codPedido)
    End If
        
End Sub

