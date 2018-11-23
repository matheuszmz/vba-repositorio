'Access e Excel

'Função Replace
'Uso: Replace("Max","x","")
'Saída: Ma
Public Function Replace(Simb, Simb1 As String, Simb2 As String)
  Dim i As Integer
     Replace = ""
     Simb = Trim(Simb)
     For i = 1 To Len(Simb)
         If Mid(Simb, i, 1) = Simb1 Then
             If Simb2 <> "" Then
                Replace = Replace & Simb2
             End If
         Else
             Replace = Replace & Mid(Simb, i, 1)
         End If
    Next
End Function

'Função que verifica se campo esta vazio ou nulo
'Uso: VazioOuNulo([B1])
'Saída: True
Public Function VazioOuNulo(x As Variant) As Boolean
    If IsNull(x) Then
        VazioOuNulo = True
    ElseIf IsEmpty(x) Then
        VazioOuNulo = True
    ElseIf x = "" Then
        VazioOuNulo = True
    Else
        VazioOuNulo = False
    End If
End Function

'Função para adicionar um item a um listBox
'Uso: AddItemList(<ListBox>, <string item>)
Function AddItemList(ctrlListBox As ListBox, ByVal strItem As String)
    ctrlListBox.AddItem Item:=strItem
End Function

'Função que Verifica se CPF é válido retornando True para válido e False para inválido
Public Function VerificarCPF(sCPF As String) As String
  Dim d1 As Integer
  Dim d2 As Integer
  Dim d3 As Integer
  Dim d4 As Integer
  Dim d5 As Integer
  Dim d6 As Integer
  Dim d7 As Integer
  Dim d8 As Integer
  Dim d9 As Integer
  Dim d10 As Integer
  Dim d11 As Integer
  Dim DV1 As Integer
  Dim DV2 As Integer
  Dim UltDig As Integer
  'Completa com zeros à esquerda caso não esteja com os 11 digitos
  If Len(sCPF) < 11 Then
    sCPF = String(11 - Len(sCPF), "0") & sCPF
  End If
  'Pega a posição do último dígito
  UltDig = Len(sCPF)
  'Sai da função caso a célula esteja vazia
  If sCPF = "00000000000" Then
    VerificarCPF = ""
    Exit Function
  End If
  'Pega cada dígito do CPF informado e coloca nas variáveis específicas
  d1 = CInt(Mid(sCPF, UltDig - 10, 1))
  d2 = CInt(Mid(sCPF, UltDig - 9, 1))
  d3 = CInt(Mid(sCPF, UltDig - 8, 1))
  d4 = CInt(Mid(sCPF, UltDig - 7, 1))
  d5 = CInt(Mid(sCPF, UltDig - 6, 1))
  d6 = CInt(Mid(sCPF, UltDig - 5, 1))
  d7 = CInt(Mid(sCPF, UltDig - 4, 1))
  d8 = CInt(Mid(sCPF, UltDig - 3, 1))
  d9 = CInt(Mid(sCPF, UltDig - 2, 1))
  d10 = CInt(Mid(sCPF, UltDig - 1, 1))    '<----- Aqui são os DVs informados
  d11 = CInt(Mid(sCPF, UltDig, 1))    '<----- no CPF analizado
  '----------- Aqui é executado o calculo para obter os digitos verificadores corretos
  DV1 = d1 + (d2 * 2) + (d3 * 3) + (d4 * 4) + (d5 * 5) + (d6 * 6) + (d7 * 7) + (d8 * 8) + (d9 * 9)
  DV1 = DV1 Mod 11    'Obtem o resto
  'se o resto for igual a 10 altera pra 0
  If DV1 = 10 Then
    DV1 = 0
  End If
  DV2 = d2 + (d3 * 2) + (d4 * 3) + (d5 * 4) + (d6 * 5) + (d7 * 6) + (d8 * 7) + (d9 * 8) + (DV1 * 9)
  DV2 = DV2 Mod 11    'Obtem o resto
  'se o resto for igual a 10 altera pra 0
  If DV2 = 10 Then
    DV2 = 0
  End If
  '---------- Fazendo a comparação dos dvs informados -------
  If d10 = DV1 And d11 = DV2 Then
    VerificarCPF = True
  Else
    VerificarCPF = False
  End If
End Function
