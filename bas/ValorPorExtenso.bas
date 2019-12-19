Attribute VB_Name = "ValorPorExtenso"
Option Explicit     'For�a a declara��o expl�cita de vari�veis
'================================================================
' FUN��O:  Extenso
' OBJETO:  Recebe um n�mero e o transforma em texto
' Fun��o principal de um conjunto de quatro rotinas
'================================================================
Public Function EXTENSO(NumValor As Currency) As String
  '--- Declara��o de vari�veis locais
  ReDim Bloco(9) As String       'Matriz: string de 1 bloco de 3 d�gitos
  ReDim TxtBloco(9) As String    'Matriz: texto para mil, milh�o, sing.
  ReDim TxtBlocoP(9) As String   'Matriz: texto para mil, milh�o, plural
  ReDim Acumula(9) As String
  Dim CmpCruz As Integer         'Compr. da string do valor (parte inteira)
  Dim EXT As String, txtValor As String
  Dim PosPtoDec As Integer, Cruzeiros As String
  Dim Cents As Variant, TotalBlocos As Integer
  Dim n As Integer, RCmpCruz As Integer
  Dim ConvBloco As String, TotalCents As String
  Dim PrimCruz As String, TxtInt$
  'Dim ContaBloco As Integer

  'Encerra fun��o se valor � zero ou branco
  If NumValor = 0 Or NumValor = Null Then Exit Function
   
  ' Define os nomes para mil, milh�o, bilh�o, etc.,
  ' no singular e no plural

  TxtBloco(2) = " mil "
  TxtBloco(3) = " milh�o "
  TxtBloco(4) = " bilh�o "
  TxtBloco(5) = " trilh�o "
  
  TxtBlocoP(2) = " mil e "
  TxtBlocoP(3) = " milh�es "
  TxtBlocoP(4) = " bilh�es "
  TxtBlocoP(5) = " trilh�es "

  EXT = ""                                'Valor tempor�rio da fun��o.
  
  txtValor = Trim(str(NumValor))          'String do valor a converter.
  PosPtoDec = InStr(Trim(txtValor), ".")  'Posi��o do ponto decimal; 0 se n�o existir
  
  Cruzeiros = Trim(Left(txtValor, IIf(PosPtoDec = 0, Len(txtValor), PosPtoDec - 1)))
  PrimCruz$ = Cruzeiros    'Reserva o valor de Reais
  CmpCruz = Len(Cruzeiros)
  Cents = Trim(Right(txtValor, IIf(PosPtoDec = 0, 0, Abs(PosPtoDec - Len(txtValor)))))
  
  'Ajusta valor de centavos ao n�vel de aproxima��o do sistema
  'Para 4, 3 e 1 decimal
  If Len(Cents) = 4 Then
    If val(Right(Cents, 2)) > 50 Then
       Cents = Format(val(Cents / 100) + 1, "00")
    Else
       Cents = Left(Cents, 2)
    End If
  End If
  
  If Len(Cents) = 3 Then
    If val(Right(Cents, 1)) > 5 Then
       Cents = Format(val(Cents / 10) + 0.1, "00")
    Else
       Cents = Left(Cents, 2)
    End If
  End If
  
  If Len(Cents) = 1 Then
     Cents = Cents & "0"
  End If

  If (CmpCruz Mod 3) = 0 Then
     TotalBlocos = (CmpCruz \ 3)
  Else
     TotalBlocos = (CmpCruz \ 3) + 1
  End If

  n% = 1
  RCmpCruz = CmpCruz      'RCmpCruz reserva valor de CmpCruz
  Do While CmpCruz > 0
     Bloco(n%) = IIf(CmpCruz > 3, Right(Cruzeiros, 3), Trim(Cruzeiros))
     Cruzeiros = IIf(CmpCruz > 3, Left(Cruzeiros, (IIf(CmpCruz < 3, 3, CmpCruz)) - 3), "")
     CmpCruz = Len(Cruzeiros)
     n% = n% + 1
  Loop

  ' Preenche matriz Acumula, que ser� usada no
  ' tratamento de exce��es
  Acumula(1) = Bloco(1)
  For n% = 2 To TotalBlocos
    Acumula(n%) = Bloco(n%) + Acumula(n% - 1)
  Next n%

  For n% = TotalBlocos To 1 Step -1     'Varre a matriz Bloco
     ' Controla plural: "milh�es", "bilh�es" etc.
     If n% > 2 And val(Bloco(n%)) > 1 Then TxtBloco(n%) = TxtBlocoP(n%)
     
     ' Controla "mil", "mil e"
     If n% = 2 Then
       If val(Bloco(1)) > 0 Then
         If (Right(Bloco(1), 2) = "00" Or val(Bloco(1)) < 100) And val(Cents) = 0 Then TxtBloco(n%) = TxtBlocoP(n%)
         If val(Bloco(n%)) = 0 Then TxtBloco(n%) = "e"
       End If
       If val(Bloco(1)) = 0 Then TxtBloco(n%) = RTrim(TxtBloco(n%))
     End If
     
     ' Adiciona "de" e "e" a "milh�es", "bilh�es"
     If n% > 2 Then
       If val(Acumula(n% - 1)) = 0 Then
         TxtBloco(n%) = TxtBloco(n%) & "de"
       Else
         If val(Cents) = 0 Then
           If val(Acumula(2)) = 0 Then
             If val(Bloco(3)) > 0 And val(Bloco(4)) > 0 Then TxtBloco(4) = TxtBloco(4) & "e "
             If val(Bloco(3)) > 0 And val(Bloco(4)) = 0 Then TxtBloco(5) = TxtBloco(5) & "e "
             If val(Bloco(3)) = 0 And val(Bloco(4)) > 0 Then TxtBloco(5) = TxtBloco(5) & "e "
           End If
           
           If val(Bloco(2)) > 0 And val(Bloco(1)) = 0 Then
            If Right(Bloco(2), 2) = "00" Or val(Bloco(2)) < 100 Then
              If val(Bloco(3)) > 0 Then TxtBloco(3) = TxtBloco(3) & "e "
              If val(Bloco(3)) = 0 And val(Bloco(4)) > 0 Then TxtBloco(4) = TxtBloco(4) & "e "
              If val(Bloco(3)) = 0 And val(Bloco(4)) = 0 Then TxtBloco(5) = TxtBloco(5) & "e "
            End If
           End If

           If val(Bloco(2)) = 0 And val(Bloco(1)) > 0 Then
            If Right(Bloco(1), 2) = "00" Or val(Bloco(1)) < 100 Then
              If val(Bloco(3)) > 0 Then TxtBloco(3) = TxtBloco(3) & "e "
              If val(Bloco(3)) = 0 And val(Bloco(4)) > 0 Then TxtBloco(4) = TxtBloco(4) & "e "
              If val(Bloco(3)) = 0 And val(Bloco(4)) = 0 Then TxtBloco(5) = TxtBloco(5) & "e "
            End If
           End If
       End If
     End If
    End If
    ConvBloco = Centena(Bloco(n%))   'Converte 1 bloco de 3 d�gitos
    
    EXT = EXT & ConvBloco            'Concatena ao valor tempor�rio da fun��o
    If ConvBloco <> "" Then EXT = EXT & TxtBloco(n%)
   Next n%
   
   TotalCents = Dezena(Cents)      'Converte centavos para texto
   If Int(NumValor) = 0 Then EXT = EXT & TotalCents & IIf(val(Cents) > 1, " centavos", " centavo")

   If Int(NumValor) = 1 Then
     If val(Cents) = 0 Then
       EXT = EXT & " real"
     Else
       EXT = EXT & " real e " & TotalCents & IIf(val(Cents) > 1, " centavos", " centavo")
     End If
   End If
   
   If Int(NumValor) > 1 Then
     If val(Cents) = 0 Then
        EXT = EXT & " reais"
     Else
        EXT = EXT & " reais e " & IIf(val(Cents) > 1, TotalCents & " centavos", TotalCents & " centavo")
     End If
   End If
   
   ' Valor final da fun��o: entre par�nteses
   EXTENSO = "( " + EXT + " )"

End Function    'Finaliza fun��o; retorna o valor por extenso

'=================================================================
' FUN��O: Centena
' Recebe parte do n�mero (entre 0 and 999) e transforma em texto
'=================================================================
Function Centena(NumText)
 Dim CT As String, x As Integer, TxtCentena As Integer
 CT = ""                       'Zera valor tempor�rio da fun��o
 If val(NumText) > 0 Then
    For x = 1 To Len(NumText)  'loop de 1 at� 3
       Select Case Len(NumText)
          Case 3:
             If val(NumText) > 99 Then
                TxtCentena = val(Left(NumText, 1))
                Select Case TxtCentena
                  Case 1
                    If Right(NumText, 2) = "00" Then
                       CT = "cem "
                    Else
                       CT = "cento "
                    End If
                  Case 2: CT = "duzentos "
                  Case 3: CT = "trezentos "
                  Case 4: CT = "quatrocentos "
                  Case 5: CT = "quinhentos "
                  Case 6: CT = "seiscentos "
                  Case 7: CT = "setecentos "
                  Case 8: CT = "oitocentos "
                  Case Else: CT = "novecentos "
                End Select
                ' Trata a exce��o: 'duzentos' e 'duzentos e'
                CT = IIf(Right(NumText, 2) > "00", CT & "e ", Left(CT, Len(CT) - 1))
                'If Right(NumText, 2) > "00" Then
                   'CT = CT & "e "
                'Else
                   'CT = Left(CT, Len(CT) - 1)
                'End If
             End If
               
             NumText = Right(NumText, 2)
          Case 2:
             CT = CT & Dezena(NumText)
             NumText = ""
          Case 1:
             CT = Unidade(NumText)
          Case Else
       End Select
    Next x
 End If
 Centena = CT  'Valor final da fun��o
End Function

'================================================================
' FUN��O: Dezena
' Recebe parte do n�mero (entre 10 and 99) e transforma em texto
'================================================================
Function Dezena(TxtDezena)
   Dim DZ As String
   Dim Unid As Integer
   DZ = ""           'anula o valor tempor�rio da fun��o
   If val(Left(TxtDezena, 1)) = 1 Then   ' Valor de 10 a 19
      Select Case val(TxtDezena)
         Case 10: DZ = "dez"
         Case 11: DZ = "onze"
         Case 12: DZ = "doze"
         Case 13: DZ = "treze"
         Case 14: DZ = "quatorze"
         Case 15: DZ = "quinze"
         Case 16: DZ = "dezesseis"
         Case 17: DZ = "dezessete"
         Case 18: DZ = "dezoito"
         Case 19: DZ = "dezenove"
         Case Else
      End Select
   Else                                 ' Valor de 20 a 99
      Select Case val(Left(TxtDezena, 1))
         Case 2: DZ = "vinte "
         Case 3: DZ = "trinta "
         Case 4: DZ = "quarenta "
         Case 5: DZ = "cinq�enta "
         Case 6: DZ = "sessenta "
         Case 7: DZ = "setenta "
         Case 8: DZ = "oitenta "
         Case 9: DZ = "noventa "
         Case Else
      End Select
      Unid = val(Right(TxtDezena, 1))
      If val(Left(TxtDezena, 1)) <> 0 Then
         If Unid <> 0 Then
            DZ = DZ & "e "
         Else
            DZ = Left(DZ, Len(DZ) - 1)
         End If
      End If
      'If Val(Left(TxtDezena, 1)) <> 0 And Unid <> 0 Then DZ = DZ & "e "
      
      DZ = DZ & Unidade(Right(TxtDezena, 1))  'Junta unidades
   End If
   Dezena = DZ                     ' Valor final da fun��o
End Function

'================================================================
' FUN��O: Unidade
' Recebe parte do n�mero (entre 1 e 9) e transforma em texto
'================================================================
Function Unidade(TxtUnidade)
   ' Atribui uma palavra a n�meros de 1 d�gito
   Select Case val(TxtUnidade)
      Case 1: Unidade = "um"
      Case 2: Unidade = "dois"
      Case 3: Unidade = "tr�s"
      Case 4: Unidade = "quatro"
      Case 5: Unidade = "cinco"
      Case 6: Unidade = "seis"
      Case 7: Unidade = "sete"
      Case 8: Unidade = "oito"
      Case 9: Unidade = "nove"
      Case Else: Unidade = ""
   End Select
End Function

Public Function Chancelamento(Inicio As Integer, Final As Integer) As String

Dim ch_X As Boolean
Dim Texto As String

ch_X = True

For Inicio = 1 To Final
    Texto = Texto + IIf(ch_X, "x", "-")
    ch_X = Not ch_X
Next

Chancelamento = Texto

End Function

