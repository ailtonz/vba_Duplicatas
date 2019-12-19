Attribute VB_Name = "ADM"
Option Compare Database
Option Explicit

Public strTabela As String

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset

Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")

If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If

rstTabela.Close


End Function

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function


Public Sub Exporta_Excel(ByVal codNotaFiscal As String)
Dim arqConfig As String
Dim arqModelo As String

'Caminho dos arquivos de Configuração e Modelo p/ preenchimento da duplicata.
arqConfig = Application.CurrentProject.Path & "\DUPLICATA.xml"
arqModelo = Application.CurrentProject.Path & "\DUPLICATA.xls"

If Dir(arqConfig, vbArchive) <> "" Then
    If Dir(arqModelo, vbArchive) <> "" Then
        'Define o arquivo de modelo como somente leitura
        SetAttr Application.CurrentProject.Path & "\DUPLICATA.xls", vbReadOnly
        
        'Objeto que dá acesso ao XML
        Dim xmldoc As MSXML2.DOMDocument
        
        'Acesso aos itens (nodes) do XML
        Dim Tipo As MSXML2.IXMLDOMNodeList
        Dim Itens As MSXML2.IXMLDOMNodeList
        
        Dim rNotaFiscal As DAO.Recordset
        
        'Variaveis de controle
        Dim Campo As String
        Dim Posicao As String
        Dim Valor As String
        Dim Duplicata As String
        Dim strEXTENSO As String
        Dim Tamanho As Integer
        Dim LimiteDoCampo As Integer
        Dim ContinuacaoDoCampo As Integer
        Dim arquivo As Variant
        Dim TotalTipo As Variant
        Dim TotalItens As Variant
        Dim x As Long
        
        Duplicata = Format(codNotaFiscal, "000000")
        
        'Instancia o objeto XMLDOM
        Set xmldoc = New MSXML2.DOMDocument
        
        Set rNotaFiscal = CurrentDb.OpenRecordset("Select * from NotasFiscais where codNotafiscal = " & codNotaFiscal)
    
        'Carrega o arquivo
        arquivo = xmldoc.Load(arqConfig)
        
        Dim objExcel As Object
        
        'Cria referencia ao EXCEL
        Set objExcel = CreateObject("Excel.Application")
        
        With objExcel
        
            .Visible = False
            .Workbooks.Open (arqModelo)
            .Sheets("DUPLICATA").Select
            .Sheets("DUPLICATA").Name = Duplicata
        
            'Coloca em ITENS da TAG chamada Tipo
            Set Tipo = PegaPorNome(xmldoc, "Tipo")
            TotalTipo = Tipo.Item(0).childNodes.Length
        
            'Retorna o Item "Item" que é o Pai dos registros detalhe
            Set Itens = PegaPorNome(xmldoc, "Item")
        
            'Retorna a quantidade de registros Detalhe
            TotalItens = Itens.Item(0).childNodes.Length
        
            If TotalTipo <> 0 Then
                For x = 0 To TotalTipo - 1
                    If TotalItens <> 0 Then
                        'Pega dentro de cada detalhe os dados necessários
                         Campo = Itens.Item(x).childNodes(0).Text
                         Posicao = Itens.Item(x).childNodes(1).Text
                         If Left(Campo, 1) = "'" Then
                            .Range(Posicao) = Campo
                         ElseIf Campo = "#EXTENSO" Then
                            Valor = Itens.Item(x).childNodes(2).Text
                            .Range(Posicao) = UCase(EXTENSO(rNotaFiscal.Fields(Valor)))
                         ElseIf Left(Campo, 1) <> "'" And Left(Campo, 1) <> "#" Then
                            .Range(Posicao) = rNotaFiscal.Fields(Campo)
                         End If
                    End If
                Next x
             End If
             
             .Range("A1").Activate
             .Visible = True
        
        End With
             
        'Descarrega da memória
        Set objExcel = Nothing
        Set rNotaFiscal = Nothing
    Else
        MsgBox "O Arquivo de modelo não existe"
    End If
Else

    MsgBox "O Arquivo de configuração não existe"

End If

End Sub

Function PegaPorNome(PXMLdoc As MSXML2.DOMDocument, SNOME As String) As MSXML2.IXMLDOMNodeList
    Set PegaPorNome = PXMLdoc.getElementsByTagName(SNOME)
End Function

