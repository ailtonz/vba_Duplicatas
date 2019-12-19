VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_FAT_NotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detalhe_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Detalhe_Format

'Banco de dados corrente
Dim dbBase As Database

'Dados da Nota Fiscal
Dim rNotaFiscal As DAO.Recordset
Dim sNotaFiscal As String

sNotaFiscal = "Select * from NotasFiscais where codNotaFiscal = " & Me.codNotaFiscal

Set dbBase = CurrentDb
Set rNotaFiscal = dbBase.OpenRecordset(sNotaFiscal)

'#########################
'### FORMATAÇÃO DA NF ####
'#########################

With Me
    .ScaleMode = 6
    .FontName = "Darf"
    .FontBold = True
    .FontSize = 9
End With

Printer.TopMargin = 0
Printer.BottomMargin = 0
Printer.LeftMargin = 110
Printer.RightMargin = 0


'################
'### OPERAÇÃO ###
'################

If rNotaFiscal.Fields("codOperacao") = "Saída" Then
  Prt 2, 124, "X"  ' Entrada
ElseIf rNotaFiscal.Fields("codOperacao") = "Entrada" Then
  Prt 2, 148, "X" ' Saida
End If

'############################
'### NATUREZA DA OPERAÇÃO ###
'############################

Prt 18, 8, Mid(rNotaFiscal.Fields("NaturezaDeOperacao"), 1, 20)
Prt 18, 72, rNotaFiscal.Fields("codCFOP")
Prt 18, 85, rNotaFiscal.Fields("InscricaoEstadualSubsTributarios")
Prt 18, 176, rNotaFiscal.Fields("DataLimiteEmissao")

'##################################
'### RAZÃO / CNPJ / DT. EMISSÃO ###
'##################################

Prt 30, 8, UCase(rNotaFiscal.Fields("Cliente"))
Prt 30, 133, rNotaFiscal.Fields("CNPJ_CPF")
Prt 30, 176, rNotaFiscal.Fields("DataDeEmissao")

'##########################
'### END / BAIRRO / CEP ###
'##########################

Prt 38, 8, rNotaFiscal.Fields("Endereco")
Prt 39, 95, rNotaFiscal.Fields("Bairro")
Prt 39, 140, rNotaFiscal.Fields("CEP")
Prt 39, 176, rNotaFiscal.Fields("DataDeSaida")

'############################
'### MUN / FONE / UF / IE ###
'############################

Prt 45, 8, rNotaFiscal.Fields("Municipio")
Prt 45, 50, rNotaFiscal.Fields("TelefoneFax")
Prt 45, 111, rNotaFiscal.Fields("Estado")
Prt 45, 133, rNotaFiscal.Fields("RG_IE")
Prt 45, 176, rNotaFiscal.Fields("HoraDeSaida")

'########################
'###      FATURAS     ###
'########################

Prt 57, 8, rNotaFiscal.Fields("Fatura_Emissao")
Prt 57, 35, rNotaFiscal.Fields("Fatura_NF")
Prt 57, 113, rNotaFiscal.Fields("Fatura_NDPL")
Prt 57, 142, rNotaFiscal.Fields("Fatura_Vencimento")
Prt 65, 142, rNotaFiscal.Fields("Fatura_Pedido")
Prt 65, 8, rNotaFiscal.Fields("Fatura_End_Cobranca")

If rNotaFiscal.Fields("Fatura_Valor") > 0 Then
    Prt 57, 64, FormatCurrency(Format(rNotaFiscal.Fields("Fatura_Valor"), "##,##0.00"))
    Prt 73, 8, UCase(EXTENSO(rNotaFiscal.Fields("Fatura_Valor")))
End If

'##########################
'### DADOS DOS PRODUTOS ###
'##########################

Prt 90, 4, rNotaFiscal.Fields("SIT_TRIB_01")
Prt 90, 15, rNotaFiscal.Fields("COD_PROD_01")
Prt 90, 35, rNotaFiscal.Fields("DESC_PROD_01")
Prt 90, 107, rNotaFiscal.Fields("UND_PROD_01")
Prt 90, 115, rNotaFiscal.Fields("QTD_PROD_01")
Prt 90, 125, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_UND_01"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_UND_01"), "##,##0.00")
Prt 90, 150, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_TOT_01"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_TOT_01"), "##,##0.00")
Prt 90, 174, IIf(IsNull(rNotaFiscal.Fields("ICMS_PROD_01")), "", rNotaFiscal.Fields("ICMS_PROD_01"))
Prt 90, 183, IIf(IsNull(rNotaFiscal.Fields("PROD_IND_01")), "", rNotaFiscal.Fields("PROD_IND_01"))
Prt 90, 190, IIf(IsNull(rNotaFiscal.Fields("VAL_PROD_01")), "", Format(rNotaFiscal.Fields("VAL_PROD_01"), "##,##0.00"))

Prt 98, 4, rNotaFiscal.Fields("SIT_TRIB_02")
Prt 98, 15, rNotaFiscal.Fields("COD_PROD_02")
Prt 98, 35, rNotaFiscal.Fields("DESC_PROD_02")
Prt 98, 107, rNotaFiscal.Fields("UND_PROD_02")
Prt 98, 115, rNotaFiscal.Fields("QTD_PROD_02")
Prt 98, 125, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_UND_02"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_UND_02"), "##,##0.00")
Prt 98, 150, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_TOT_02"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_TOT_02"), "##,##0.00")
Prt 98, 174, IIf(IsNull(rNotaFiscal.Fields("ICMS_PROD_02")), "", rNotaFiscal.Fields("ICMS_PROD_02"))
Prt 98, 183, IIf(IsNull(rNotaFiscal.Fields("PROD_IND_02")), "", rNotaFiscal.Fields("PROD_IND_02"))
Prt 98, 190, IIf(IsNull(rNotaFiscal.Fields("VAL_PROD_02")), "", Format(rNotaFiscal.Fields("VAL_PROD_02"), "##,##0.00"))

Prt 106, 4, rNotaFiscal.Fields("SIT_TRIB_03")
Prt 106, 15, rNotaFiscal.Fields("COD_PROD_03")
Prt 106, 35, rNotaFiscal.Fields("DESC_PROD_03")
Prt 106, 107, rNotaFiscal.Fields("UND_PROD_03")
Prt 106, 115, rNotaFiscal.Fields("QTD_PROD_03")
Prt 106, 125, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_UND_03"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_UND_03"), "##,##0.00")
Prt 106, 150, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_TOT_03"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_TOT_03"), "##,##0.00")
Prt 106, 174, IIf(IsNull(rNotaFiscal.Fields("ICMS_PROD_03")), "", rNotaFiscal.Fields("ICMS_PROD_03"))
Prt 106, 183, IIf(IsNull(rNotaFiscal.Fields("PROD_IND_03")), "", rNotaFiscal.Fields("PROD_IND_03"))
Prt 106, 190, IIf(IsNull(rNotaFiscal.Fields("VAL_PROD_03")), "", Format(rNotaFiscal.Fields("VAL_PROD_03"), "##,##0.00"))

Prt 114, 4, rNotaFiscal.Fields("SIT_TRIB_04")
Prt 114, 15, rNotaFiscal.Fields("COD_PROD_04")
Prt 114, 35, rNotaFiscal.Fields("DESC_PROD_04")
Prt 114, 107, rNotaFiscal.Fields("UND_PROD_04")
Prt 114, 115, rNotaFiscal.Fields("QTD_PROD_04")
Prt 114, 125, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_UND_04"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_UND_04"), "##,##0.00")
Prt 114, 150, Space(12 - Len(Format(rNotaFiscal.Fields("PRC_TOT_04"), "##,##0.00"))) & Format(rNotaFiscal.Fields("PRC_TOT_04"), "##,##0.00")
Prt 114, 174, IIf(IsNull(rNotaFiscal.Fields("ICMS_PROD_04")), "", rNotaFiscal.Fields("ICMS_PROD_04"))
Prt 114, 183, IIf(IsNull(rNotaFiscal.Fields("PROD_IND_04")), "", rNotaFiscal.Fields("PROD_IND_04"))
Prt 114, 190, IIf(IsNull(rNotaFiscal.Fields("VAL_PROD_04")), "", Format(rNotaFiscal.Fields("VAL_PROD_04"), "##,##0.00"))

'###############################
'### DESCRICAO DOS SERVICOS ####
'###############################

Prt 130, 8, rNotaFiscal.Fields("DESC_SERV_01")
Prt 134, 8, rNotaFiscal.Fields("DESC_SERV_02")
Prt 138, 8, rNotaFiscal.Fields("DESC_SERV_03")
Prt 142, 8, rNotaFiscal.Fields("DESC_SERV_04")
Prt 148, 8, rNotaFiscal.Fields("DESC_SERV_05")

If rNotaFiscal.Fields("RET_IRRF_VAL") > 0 Then
    Prt 150, 80, "RETENCAO IRRF            : " & rNotaFiscal.Fields("RET_IRRF") & "% = " & FormatCurrency(Format(rNotaFiscal.Fields("RET_IRRF_VAL"), "##,##0.00"))
End If

If rNotaFiscal.Fields("RET_PIS_COFINS_CSLL") > 0 Then
    Prt 154, 80, "RETENCAO PIS / COFINS / CSLL : " & rNotaFiscal.Fields("RET_PIS_COFINS_CSLL") & "% = " & FormatCurrency(Format(rNotaFiscal.Fields("RET_PIS_COFINS_CSLL_VAL"), "##,##0.00"))
End If

If rNotaFiscal.Fields("VAL_ISS") > 0 And rNotaFiscal.Fields("PCT_ISS") > 0 Then
    Prt 144, 173, rNotaFiscal.Fields("PCT_ISS")
    Prt 144, 183, Format(rNotaFiscal.Fields("VAL_ISS"), "##,##0.00")
End If

If rNotaFiscal.Fields("TOT_SERV") > 0 Then
    Prt 154, 174, FormatCurrency(Format(rNotaFiscal.Fields("TOT_SERV"), "##,##0.00"))
End If

'###########################
'### CALCULOS DO IMPOSTO ###
'###########################

If Not IsNull(rNotaFiscal.Fields("BaseDeCalculoDoICMS")) Then
    Prt 170, 10, IIf(rNotaFiscal.Fields("BaseDeCalculoDoICMS") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("BaseDeCalculoDoICMS"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoICMS")) Then
    Prt 170, 50, IIf(rNotaFiscal.Fields("ValorDoICMS") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorDoICMS"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("BaseDeCalculoICMSSubstituicao")) Then
    Prt 170, 85, IIf(rNotaFiscal.Fields("BaseDeCalculoICMSSubstituicao") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("BaseDeCalculoICMSSubstituicao"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoICMSSubstituicao")) Then
    Prt 170, 125, IIf(rNotaFiscal.Fields("ValorDoICMSSubstituicao") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorDoICMSSubstituicao"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDosProdutos")) Then
    Prt 170, 165, IIf(rNotaFiscal.Fields("ValorTotalDosProdutos") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorTotalDosProdutos"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoFrete")) Then
    Prt 178, 10, IIf(rNotaFiscal.Fields("ValorDoFrete") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorDoFrete"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoServico")) Then
    Prt 178, 50, IIf(rNotaFiscal.Fields("ValorDoServico") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorDoServico"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("OutrasDespesasAcessorias")) Then
    Prt 178, 85, IIf(rNotaFiscal.Fields("OutrasDespesasAcessorias") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("OutrasDespesasAcessorias"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDoIPI")) Then
    Prt 178, 125, IIf(rNotaFiscal.Fields("ValorTotalDoIPI") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorTotalDoIPI"), "##,##0.00")))
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDaNota")) Then
    Prt 178, 165, IIf(rNotaFiscal.Fields("ValorTotalDaNota") = 0, "----", FormatCurrency(Format(rNotaFiscal.Fields("ValorTotalDaNota"), "##,##0.00")))
End If

'###############################
'### TRANSPORTADOR / VOLUMES ###
'###############################

''TRANSPORTADOR

Prt 190, 10, UCase(rNotaFiscal.Fields("Transportadora"))
Prt 190, 107, rNotaFiscal.Fields("FretePorConta")
Prt 190, 115, rNotaFiscal.Fields("PlacaDoVeiculo")
Prt 190, 140, rNotaFiscal.Fields("UFDaPlaca")
Prt 190, 150, rNotaFiscal.Fields("TranspCNPJ_CPF")

Prt 198, 10, rNotaFiscal.Fields("TranspEndereco")
Prt 198, 85, rNotaFiscal.Fields("TranspMunicipio")
Prt 198, 140, rNotaFiscal.Fields("TranspEstado")
Prt 198, 150, rNotaFiscal.Fields("TranspRG_IE")

''VOLUMES

Prt 207, 10, rNotaFiscal.Fields("Quantidade")
Prt 207, 30, rNotaFiscal.Fields("Especie")
Prt 207, 55, rNotaFiscal.Fields("Marca")
Prt 207, 90, rNotaFiscal.Fields("NumeroDeControle")
Prt 207, 142, rNotaFiscal.Fields("PesoBruto")
Prt 207, 174, rNotaFiscal.Fields("PesoLiquido")

'########################
'### DADOS ADICIONAIS ###
'########################

Prt 233, 50, IIf(rNotaFiscal.Fields("Duplicata_A_Valor") = 0, "", Format(rNotaFiscal.Fields("Duplicata_A_Valor"), "##,##0.00"))
Prt 233, 75, rNotaFiscal.Fields("Duplicata_A_Vencimento")

Prt 238, 50, IIf(rNotaFiscal.Fields("Duplicata_B_Valor") = 0, "", Format(rNotaFiscal.Fields("Duplicata_B_Valor"), "##,##0.00"))
Prt 238, 75, rNotaFiscal.Fields("Duplicata_B_Vencimento")

Prt 243, 50, IIf(rNotaFiscal.Fields("Duplicata_C_Valor") = 0, "", Format(rNotaFiscal.Fields("Duplicata_C_Valor"), "##,##0.00"))
Prt 243, 75, rNotaFiscal.Fields("Duplicata_C_Vencimento")


Exit_Detalhe_Format:
    Exit Sub

Err_Detalhe_Format:
    MsgBox Err.Description
    Resume Exit_Detalhe_Format
    
End Sub

Sub Prt(Linha, Coluna, Texto)

Dim mSup As Long
Dim mEsq As Long

'If Me.Page > 1 Then
'   If Me.Page = 2 Then
'      mSup = 18#
'   Else
'      mSup = 17# - ((Me.Page - 1) * 0.5)
'   End If
'Else
'   mSup = 14#
'End If

mEsq = 3
mSup = 13 - 1

With Me
  .CurrentY = Linha + mSup
  .CurrentX = Coluna + mEsq
  If Not IsNull(Texto) Then
     .Print Texto
  End If
End With

End Sub

Private Sub Report_Close()
    DoCmd.Restore
End Sub

Private Sub Report_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
