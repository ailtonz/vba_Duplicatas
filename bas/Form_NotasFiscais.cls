VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnDuplicata_Click()
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Call Exporta_Excel(Me.codNotaFiscal)
End Sub

Private Sub cboCliente_Click()
    Me.txtEndereco = Me.cboCliente.Column(1)
    Me.txtBairro = Me.cboCliente.Column(2)
    Me.txtCep = Me.cboCliente.Column(3)
    Me.txtMunicipio = Me.cboCliente.Column(4)
    Me.txtEstado = Me.cboCliente.Column(5)
    Me.txtTelefoneFax = Me.cboCliente.Column(6)
    Me.txtCNPJ_CPF = Me.cboCliente.Column(7)
    Me.txtRG_IE = Me.cboCliente.Column(8)
End Sub

Private Sub cboCliente_Exit(Cancel As Integer)
    Call cboCliente_Click
End Sub

Private Sub cbocodCFOP_Click()
    Me.txtNaturezaDeOperacao = Me.cboCodCFOP.Column(1)
    Me.cboCodOperacao = Me.cboCodCFOP.Column(2)
End Sub


Private Sub btnFechar_Click()
On Error GoTo Err_btnFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_btnFechar_Click:
    Exit Sub

Err_btnFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_btnFechar_Click

End Sub

Private Sub btnSalvar_Click()
On Error GoTo Err_btnSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_btnSalvar_Click:
    Exit Sub

Err_btnSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_btnSalvar_Click

End Sub

Private Sub btnVisualizar_Click()
On Error GoTo Err_btnVisualizar_Click

    Dim stDocName As String

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    stDocName = "FAT_NotasFiscais"
    DoCmd.OpenReport stDocName, acPreview, , "codNotaFiscal = " & Me.txtCodigo

Exit_btnVisualizar_Click:
    Exit Sub

Err_btnVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_btnVisualizar_Click

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.txtCodigo = NovoCodigo(Me.RecordSource, Me.txtCodigo.ControlSource)
       Me.txtDataDeEmissao = Format(Now(), "dd/mm/yy")
    End If
    
End Sub

Private Sub Form_Close()
' Turn the MouseWheel On

Dim blRet As Boolean

    blRet = MouseWheelON
    DoCmd.Restore
End Sub

Private Sub Form_Open(Cancel As Integer)
' Turn the MouseWheel Off
' Call our MouseHook function in the MouseHook dll.
' Please not the Optional GlobalHook BOOLEAN parameter
' Several developers asked for the MouseHook to be able to work with
' multiple instances of Access. In order to accomodate this request I
' have modified the function to allow the caller to
' specify a thread specific(this current instance of Access only) or
' a global(all applications) MouseWheel Hook.
' Only use the GlobalHook if you will be running multiple instances of Access!

Dim blRet As Boolean


blRet = MouseWheelOFF(False)

End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
End Sub

Private Sub txtPRC_UND_01_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_01)
    val_02 = setString(Me.txtPRC_UND_01)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_01 = Valor
    
    somaTotal

End Sub

Private Sub txtQTD_PROD_01_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_01)
    val_02 = setString(Me.txtPRC_UND_01)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_01 = Valor
    
    somaTotal

End Sub


Private Sub txtPRC_UND_02_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_02)
    val_02 = setString(Me.txtPRC_UND_02)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_02 = Valor
    
    somaTotal

End Sub

Private Sub txtQTD_PROD_02_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_02)
    val_02 = setString(Me.txtPRC_UND_02)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_02 = Valor
    
    somaTotal

End Sub

Private Sub txtPRC_UND_03_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_03)
    val_02 = setString(Me.txtPRC_UND_03)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_03 = Valor
    
    somaTotal

End Sub

Private Sub txtQTD_PROD_03_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_03)
    val_02 = setString(Me.txtPRC_UND_03)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_03 = Valor
    
    somaTotal

End Sub

Private Sub txtPRC_UND_04_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_04)
    val_02 = setString(Me.txtPRC_UND_04)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_04 = Valor
    
    somaTotal

End Sub

Private Sub txtQTD_PROD_04_Exit(Cancel As Integer)
    Dim Valor As String
    Dim val_01, val_02 As String
    
    val_01 = setString(Me.txtQTD_PROD_04)
    val_02 = setString(Me.txtPRC_UND_04)
    Valor = SomaProd(val_01, val_02)
    Me.txtPRC_TOT_04 = Valor
    
    somaTotal

End Sub

Private Function SomaProd(ByVal val_01 As String, ByVal val_02 As String) As String
    If val_01 <> "" And val_02 <> "" Then
        SomaProd = (val_01 * val_02)
    Else
        SomaProd = ""
    End If
End Function

Private Function setString(val As Variant) As String
    If IsNull(val) Then
        setString = ""
    Else
        setString = CStr(val)
    End If
End Function

Private Sub txtPCT_ISS_Exit(Cancel As Integer)
    Me.txtVAL_ISS = (Me.txtTOT_SERV / 100) * Me.txtPCT_ISS
End Sub

Private Sub txtRET_IRRF_Exit(Cancel As Integer)
    Me.txtRET_IRRF_VAL = (Me.txtTOT_SERV / 100) * Me.txtRET_IRRF
End Sub

Private Sub txtRET_PIS_COFINS_CSLL_Exit(Cancel As Integer)
    Me.txtRET_PIS_COFINS_CSLL_VAL = (Me.txtTOT_SERV / 100) * Me.txtRET_PIS_COFINS_CSLL
End Sub

Private Function somaTotal()
    
    somaTotal = IIf(setString(Me.txtPRC_TOT_01) <> "", Me.txtPRC_TOT_01, 0) + IIf(setString(Me.txtPRC_TOT_02) <> "", Me.txtPRC_TOT_02, 0) + IIf(setString(Me.txtPRC_TOT_03) <> "", Me.txtPRC_TOT_03, 0) + IIf(setString(Me.txtPRC_TOT_04) <> "", Me.txtPRC_TOT_04, 0)
    Me.txtValorTotalDosProdutos = somaTotal
    Me.txtValorTotalDaNota = Me.txtValorTotalDosProdutos
    CarregarFatura
    
End Function

Private Sub txtTOT_SERV_Exit(Cancel As Integer)
Dim valor1 As Long
Dim valor2 As Long
'Dim valor3 As Long
Dim IRRF As String
Dim PIS_COFINS_CSLL As String
Dim PCT_ISS As String

valor1 = "700,00"
valor2 = "4.999,99"
'valor3 = "5.000,00"

PCT_ISS = "5"
IRRF = "1,5"
PIS_COFINS_CSLL = "4,65"

If Me.txtTOT_SERV = 0 Then
    'No Caso de não haver valor de Serviços estes campos são zerados
    Me.txtPCT_ISS = 0
    Me.txtVAL_ISS = 0
    Me.txtRET_IRRF = 0
    Me.txtRET_IRRF_VAL = 0
    Me.txtRET_PIS_COFINS_CSLL = 0
    Me.txtRET_PIS_COFINS_CSLL_VAL = 0
    Me.txtValorTotalDaNota = 0
    CarregarFatura
ElseIf Me.txtTOT_SERV > valor1 And Me.txtTOT_SERV < valor2 Then
    'Zera estes campos
    Me.txtPCT_ISS = 0
    Me.txtVAL_ISS = 0
    Me.txtRET_IRRF = 0
    Me.txtRET_IRRF_VAL = 0
    Me.txtRET_PIS_COFINS_CSLL = 0
    Me.txtRET_PIS_COFINS_CSLL_VAL = 0
    'Preenche o Percentual do IRRF
    Me.txtRET_IRRF = IRRF
    'Calcula o valor do IRRF
    Call txtRET_IRRF_Exit(0)
    'O Total da NF é igual ao Total do Serviço
    Me.txtValorTotalDaNota = Me.txtTOT_SERV
    'Preenche os campos da Fatura (Emissão,Numero da NF e Valor da Fatura)
    CarregarFatura
ElseIf Me.txtTOT_SERV >= valor2 Then
    'Zera estes campos
    Me.txtPCT_ISS = 0
    Me.txtVAL_ISS = 0
    Me.txtRET_IRRF = 0
    Me.txtRET_IRRF_VAL = 0
    Me.txtRET_PIS_COFINS_CSLL = 0
    Me.txtRET_PIS_COFINS_CSLL_VAL = 0
    'Preenche o Percentual do ISS
    Me.txtPCT_ISS = PCT_ISS
    'Calcula o valor do ISS
    Call txtPCT_ISS_Exit(0)
    'Preenche o Percentual do IRRF
    Me.txtRET_IRRF = IRRF
    'Calcula o valor do IRRF
    Call txtRET_IRRF_Exit(0)
    'Preenche o Percentual do PIS,COFINS e CSLL
    Me.txtRET_PIS_COFINS_CSLL = PIS_COFINS_CSLL
    'Calcula o valor do PIS,COFINS e CSLL
    Call txtRET_PIS_COFINS_CSLL_Exit(0)
    'O Total da NF é igual ao Total do Serviço
    Me.txtValorTotalDaNota = Me.txtTOT_SERV
    'Preenche os campos da Fatura (Emissão,Numero da NF e Valor da Fatura)
    CarregarFatura
Else
    'Zera estes campos
    Me.txtPCT_ISS = 0
    Me.txtVAL_ISS = 0
    Me.txtRET_IRRF = 0
    Me.txtRET_IRRF_VAL = 0
    Me.txtRET_PIS_COFINS_CSLL = 0
    Me.txtRET_PIS_COFINS_CSLL_VAL = 0
    'O Total da NF é igual ao Total do Serviço
    Me.txtValorTotalDaNota = Me.txtTOT_SERV
    'Preenche os campos da Fatura (Emissão,Numero da NF e Valor da Fatura)
    CarregarFatura
End If
    
End Sub

Private Sub CarregarFatura()
'---------------------------------------------
'Responsável por preencher os campos(Fatura):
'   > Emissão
'   > Num. NF
'   > Valor
'---------------------------------------------
Dim Limite1 As Currency
Dim Limite2 As Currency

Limite1 = "5.999,99"
Limite2 = "6.000,00"

    Me.txtFatura_Emissao = Me.txtDataDeEmissao
    Me.txtFatura_NF = Format(Me.txtCodigo, "000000")
    Me.txtFatura_Valor = Me.txtValorTotalDaNota

    If Me.txtFatura_Valor <= Limite1 Then
        Me.txtFatura_Valor = Me.txtFatura_Valor - Me.txtRET_IRRF_VAL
    ElseIf Me.txtFatura_Valor >= Limite2 Then
        Me.txtFatura_Valor = Me.txtFatura_Valor - Me.txtRET_IRRF_VAL - Me.txtRET_PIS_COFINS_CSLL_VAL
    End If

End Sub
