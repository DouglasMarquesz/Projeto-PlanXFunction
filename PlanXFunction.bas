Attribute VB_Name = "PlanXFunction"
' MÓDULO CRIADO EM 13/09/2022 POR DOUGLAS JR MARQUES ANGELO
'
' v1.2.7
'
'Funções:

'Deletar todos os dados da planilha menos o cabeçalho
Function deletarTodosDados(planilha As String)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
Dim linhas As Integer
    linhas = ActiveSheet.UsedRange.Rows.Count
    If linhas > 1 Then
        ActiveSheet.Rows("2:" & linhas).Delete
    End If
End Function

'Alterar dados
Function alterarDados(planilha As String, TextoAlvo As String, coluna As String, NovoTexto As String)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
With Worksheets(planilha).Range(UCase(coluna) + ":" + UCase(coluna))
    Set c = .Find(TextoAlvo, LookIn:=xlValues, lookat:=xlWhole)
    If Not c Is Nothing Then
        c.Select
        ActiveCell.Value = NovoTexto
        alterarDados = True
    Else
        alterarDados = False
    End If
End With
End Function

'Analisar se existe
Function analisarExiste(planilha As String, TextoAlvo As String, coluna As String)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
With Worksheets(planilha).Range(UCase(coluna) + ":" + UCase(coluna))
    Set c = .Find(TextoAlvo, LookIn:=xlValues, lookat:=xlWhole)
    If Not c Is Nothing Then
        c.Select
        analisarExiste = True
    Else
        analisarExiste = False
    End If
End With
End Function

'Analisar se existe com uma condição específica
Function analisarExisteCond(planilha As String, TextoAlvo As String, coluna As String, TextoAlvo2 As String, puloCélula As Integer)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
n = Range(coluna + "1").End(xlDown).Row
n = Range(coluna & Cells.Rows.Count).End(xlUp).Row
Range(coluna + "1").Select
For i = 1 To n
    Range(coluna + CStr(i)).Select
    If ActiveCell.Value = TextoAlvo Then
        If ActiveCell.Offset(0, puloCélula).Value = TextoAlvo2 Then
            c.Select
            analisarExisteCond = True
            Exit Function
        Else
            analisarExisteCond = False
        End If
    Else
        analisarExisteCond = False
    End If
Next i

End Function

'Quantidade de vezes que existe
Function contarExiste(planilha As String, TextoAlvo As String, coluna As String)
On Error Resume Next
Dim qtd As Integer
qtd = 0
Worksheets(planilha).Activate
Worksheets(planilha).Select
n = Range(coluna + "1").End(xlDown).Row
n = Range(coluna & Cells.Rows.Count).End(xlUp).Row
Range(coluna + "1").Select
For i = 1 To n
    Range(coluna + CStr(i)).Select
    If ActiveCell.Value = TextoAlvo Then
        qtd = qtd + 1
    End If
Next i
contarExiste = qtd
End Function

'Quantidade de vezes que existe com condição
Function contarExisteCond(planilha As String, TextoAlvo As String, coluna As String, TextoAlvo2 As String, puloCélula As Integer)
On Error Resume Next
Dim qtd As Integer
qtd = 0
Worksheets(planilha).Activate
Worksheets(planilha).Select
n = Range(coluna + "1").End(xlDown).Row
n = Range(coluna & Cells.Rows.Count).End(xlUp).Row
Range(coluna + "1").Select
For i = 1 To n
    Range(coluna + CStr(i)).Select
    If (ActiveCell.Value = texto And ActiveCell.Offset(0, puloCélula).Value = condicaoTexto) Then
        qtd = qtd + 1
    End If
Next i
contarExisteCond = qtd
End Function

'Inserir dados em uma planilha com ou sem contador (ATUALIZADO COM ARRAY)
Function inserirDados(planilha As String, dados As Variant, indexCodigo As Boolean)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
Dim cod As Integer
n = Range("A1").End(xlDown).Row
n = Range("A" & Cells.Rows.Count).End(xlUp).Row
Range("A" + CStr(n + 1)).Select

If (indexCodigo = True) Then
    If (IsNumeric(ActiveCell.Offset(-1, 0).Value)) = False Then
        ActiveCell.Value = 0
    Else
        cod = ActiveCell.Offset(-1, 0) + 1
    End If
    ActiveCell.Offset(0, 0).Value = cod
    For i = 1 To (UBound(dados) + 1)
        ActiveCell.Offset(0, i).Value = dados(i - 1)
    Next i
Else
    For i = 0 To (UBound(dados) + 1) - 1
        ActiveCell.Offset(0, i).Value = dados(i)
    Next i
End If
End Function

'Retornar quantidade de linhas existentes (ultima linha da planilha)
Function ultimaLinhaCod(planilha As String)
Worksheets(planilha).Activate
Worksheets(planilha).Select
n = Range("A1").End(xlDown).Row
n = Range("A" & Cells.Rows.Count).End(xlUp).Row
ultimaLinhaCod = n
End Function

'Deletar dados da planilha
Function deletarDados(planilha As String, TextoAlvo As String, coluna As String, duplicados As Boolean)
deletarDados = False
inicioDelete:
Worksheets(planilha).Activate
Worksheets(planilha).Select
With Worksheets(planilha).Range(UCase(coluna) + ":" + UCase(coluna))
    Set c = .Find(TextoAlvo, LookIn:=xlValues, lookat:=xlWhole)
    If Not c Is Nothing Then
        c.Activate
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.EntireRow.Delete
        deletarDados = True
        If duplicados = False Then
            Exit Function
        Else
            GoTo inicioDelete
        End If
    Else
        If deletarDados <> True Then
            deletarDados = False
        End If
    End If
End With
End Function

'Deletar dados da planilha com condição
Function deletarDadosCond(planilha As String, TextoAlvo As String, coluna As String, TextoAlvo2 As String, puloCélula As Integer, duplicados As Boolean)
On Error Resume Next
deletarDadosCond = False
inicioDeleteCond:
Worksheets(planilha).Activate
Worksheets(planilha).Select
n = Range(coluna + "1").End(xlDown).Row
n = Range(coluna & Cells.Rows.Count).End(xlUp).Row
Range(coluna + "1").Select
For i = 1 To n
    Range(coluna + CStr(i)).Select
    If ActiveCell.Value = texto Then
        If ActiveCell.Offset(0, puloCélula).Value = condicaoTexto Then
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.EntireRow.Delete
            deletarDadosCond = True
            If duplicados = False Then
                Exit Function
            End If
            GoTo inicioDeleteCond
        Else
            If deletarDadosCond <> True Then
                deletarDadosCond = False
            End If
        End If
    Else
        If deletarDadosCond <> True Then
            deletarDadosCond = False
        End If
    End If
Next i
End Function

'Inserir Cabeçalho no ListView (Atualização 09/11/2022)
Function listViewTop(listview As Object, dados As Variant, celulaTamanho As Variant)
With listview
    .ColumnHeaders.Clear
    .View = lvwReport
    .FullRowSelect = True
    .Gridlines = True
    For i = 0 To UBound(dados)
        If dados(i) <> "" Then
            .ColumnHeaders.Add , , dados(i), celulaTamanho(i)
        End If
    Next i
End With
End Function

'Inserir Toda planilha selecionado no ListView (Atualização 09/11/2022)
Function listViewDados(planilha As String, listview As Object, qtdColunas As Integer, cabeçalho As Boolean)
On Error Resume Next
Worksheets(planilha).Activate
Worksheets(planilha).Select
Range("A1").Select
n = Range("A1").End(xlDown).Row
n = Range("A" & Cells.Rows.Count).End(xlUp).Row
If cabeçalho = False Then
    X = 1
    Else
    X = 0
End If
For i = 0 To n - (X + 1)
Set itmx = listview.ListItems.Add(, , ActiveCell.Offset(i + X, 0).Value)
    For j = 1 To qtdColunas - 1
        itmx.SubItems(j) = CStr(ActiveCell.Offset(i + X, j).Value)
    Next j
Next i
End Function

'Inserir filtro em edtTexto para digitalização (Atualização 09/11/2022)
Function listViewFilter(ListViewSecundario As Object, listviewFinal As Object, inputbox As Object, qtdColunas As Integer, filtroColuna As Integer)
On Error Resume Next
Dim palavra As String
Dim objetivo, texto As String
Range("A1").Select
listviewFinal.ListItems.Clear
texto = UCase(inputbox.Text)
For i = 1 To ListViewSecundario.ListItems.Count
    If filtroColuna = 0 Then
        objetivo = UCase(ListViewSecundario.ListItems.Item(i))
    Else
        objetivo = UCase(ListViewSecundario.ListItems.Item(i).SubItems(filtroColuna))
    End If
    If InStr(objetivo, texto) = 1 Then
        Set itmx = listviewFinal.ListItems.Add(, , ListViewSecundario.ListItems.Item(i))
        For j = 1 To qtdColunas - 1
            itmx.SubItems(j) = CStr(ListViewSecundario.ListItems.Item(i).SubItems(j))
        Next j
    End If
Next i
End Function

'Buscar dados em toda a planilha (Sem repetições)
Function buscarDados(planilha As String, valor As String, coluna As String, qtdItens As Integer, resultado As Variant)
Worksheets(planilha).Activate
Worksheets(planilha).Select
With Worksheets(planilha).Range(UCase(coluna) + ":" + UCase(coluna))
    Set c = .Find(valor, LookIn:=xlValues, lookat:=xlWhole)
    If Not c Is Nothing Then
        c.Activate
        For i = 0 To qtdItens - 1
            resultado(i) = ActiveCell.Offset(0, i).Value
        Next i
        buscarDados = True
    Else
        buscarDados = False
    End If
End With
End Function

'FUNÇÃO NOVA
'Buscar dados em toda a planilha (Sem repetições) E PEGAR DO PRIMEIRO
Function buscarDadosInicio(planilha As String, valor As String, coluna As String, qtdItens As Integer, resultado As Variant)
Worksheets(planilha).Activate
Worksheets(planilha).Select
With Worksheets(planilha).Range(UCase(coluna) + ":" + UCase(coluna))
    Set c = .Find(valor, LookIn:=xlValues, lookat:=xlWhole)
    If Not c Is Nothing Then
        c.Activate
        celula = "A" + CStr(ActiveCell.Row)
        Range(celula).Select
        For i = 0 To qtdItens - 1
            resultado(i) = ActiveCell.Offset(0, i).Value
            
        Next i
        buscarDadosInicio = True
    Else
        buscarDadosInicio = False
    End If
End With
End Function


'ATUALIZAÇÃO NOVA FUNÇÃO
'Buscar dados em toda a planilha com condição (Sem repetições)
Function buscarDadosCond(planilha As String, valor As String, coluna As String, qtdItens As Integer, resultado() As String, condicaoTexto As String, puloCélula As Integer)
Worksheets(planilha).Activate
Worksheets(planilha).Select
qtd = 0
n = Range(coluna + "1").End(xlDown).Row
n = Range(coluna & Cells.Rows.Count).End(xlUp).Row
Range(coluna + "1").Select
For i = 1 To n
    Range(coluna + CStr(i)).Select
    If (ActiveCell.Value = valor And ActiveCell.Offset(0, puloCélula).Value = condicaoTexto) Then
        Range(Selection, Selection.End(xlToRight)).Select
        For j = 0 To qtdItens - 1
            resultado(j) = ActiveCell.Offset(0, j).Value
        Next j
        Exit Function
    End If
Next i
End Function

'ATUALIZAÇÃO NOVA FUNÇÃO
'Gerador de números aleatórios
Function numeroAleatorio(intervalo As Integer)
Randomize
numeroAleatorio = CInt(Int((intervalo * Rnd()) + 1))
End Function

'Função de finalização
Function sair(formulario As Object, Confirmar As Boolean)
Dim wb As Workbook
If Confirmar = False Then
    If Application.Workbooks.Count = 1 Then
        ActiveWorkbook.SAVE
        Application.Quit
    Else
        Application.Visible = True
        For Each wb In Workbooks
            If wb.Name = ThisWorkbook.Name Then
                wb.Close SaveChanges:=True
                Exit For
            End If
        Next wb
    End If
Else
    If MsgBox("Tem certeza que deseja finalizar?", vbInformation + vbYesNo, "Finalizar") = vbYes Then
        If Application.Workbooks.Count = 1 Then
            ActiveWorkbook.SAVE
            Application.Quit
        Else
            Application.Visible = True
            For Each wb In Workbooks
                If wb.Name = ThisWorkbook.Name Then
                    wb.Close SaveChanges:=True
                    Exit For
                End If
            Next wb
        End If
    End If
End If
End Function

'Diretório Raiz de onde o projeto está aberto
Function diretorioRaiz()
    On Error Resume Next
    diretorioRaiz = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
End Function

'Retornar data de modificação do arquivo
Function dataModificacao(ByVal diretorio As String) As Date
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.GetFile(diretorio)
    dataModificacao = file.DateLastModified
End Function

'Função verificar CPF
Function verificarCPF(CPF As String) As Boolean
Dim cpf2 As String
    'Remover Pontuação
    cpf2 = Replace(CPF, ".", "")
    CPF = Replace(cpf2, "-", "")
    Dim strCPF As String
    Dim numDV1 As Integer
    Dim numDV2 As Integer
    Dim numCheckDV1 As Integer
    Dim numCheckDV2 As Integer
    Dim i As Integer
    If CPF = "" Then
        verificarCPF = False
        Exit Function
    End If
    'Módulo 11 do 1° digito'
    strCPF = Right$(String$(11, "0") + CPF, 11)
    numCheckDV1 = Val(Mid$(strCPF, 10, 1))
    numCheckDV2 = Val(Mid$(strCPF, 11, 1))
    For i = 1 To 9
        numDV1 = numDV1 + Val(Mid$(strCPF, i, 1)) * i
    Next i
    numDV1 = numDV1 Mod 11
    If numDV1 = 10 Then numDV1 = 0
    
    'Modulo 11 - 2° digito verificador'
    If numDV1 <> numCheckDV1 Then Exit Function
    For i = 2 To 10
        numDV2 = numDV2 + Val(Mid$(strCPF, i, 1)) * (i - 1)
    Next i
    numDV2 = numDV2 Mod 11
    If numDV2 = 10 Then numDV2 = 0
    If numDV2 <> numCheckDV2 Then Exit Function
        verificarCPF = True
End Function

'Mensagem erro (formulário frmErrorU)
Function ErrorForm(numero As String, mensagem As String)
    frmErrorU.NumeroError.Caption = numero
    frmErrorU.DescricaoError.Caption = mensagem
    frmErrorU.Show
End Function
