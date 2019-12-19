VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RelatoriosCriterios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Descricao_Click()
    Me.Valor = Me.Descricao.Column(0)
End Sub

Private Sub Descricao_Enter()
    Me.Descricao.RowSource = Me.OrigemDaLinha.Value
    Me.Descricao.ColumnWidths = Me.LarguraDasColunas
End Sub


