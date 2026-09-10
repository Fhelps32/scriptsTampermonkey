Attribute VB_Name = "FormatarDiagnosticoUGB"
'==============================================================================
' Formatar diagnóstico dos bancos (UGB)
'
' Deixa apresentável o CSV gerado pelo botão "Baixar CSV" do script
' "Importar bancos de questões": título, cabeçalho fixo, filtro, cores por
' situação, linha de totais e área de impressão.
'
' COMO INSTALAR (uma vez só)
'   1. No Excel, Alt+F11 abre o editor do VBA.
'   2. Menu Arquivo > Importar Arquivo… e escolha este .bas.
'      Se preferir colar (Inserir > Módulo), NÃO copie a primeira linha
'      "Attribute VB_Name": ela só vale na importação e o editor recusa.
'   3. Feche o editor. A macro fica na pasta PESSOAL.XLSB se você a
'      importar lá — assim vale para qualquer diagnóstico futuro.
'
' COMO USAR
'   1. Abra o CSV do diagnóstico.
'   2. Alt+F8 > FormatarDiagnosticoBancos > Executar.
'   3. No fim ela oferece salvar como .xlsx — aceite, senão o Excel joga
'      a formatação fora ao salvar (CSV não guarda formatação).
'
' Pode rodar de novo na mesma planilha: ela refaz a formatação do zero.
' As colunas são localizadas pelo NOME no cabeçalho, então continua
' funcionando se o script passar a exportar outras categorias além de
' AV1/AV2.
'==============================================================================
Option Explicit

'--- paleta (as mesmas cores do painel do script) -----------------------------
Private Const COR_TITULO As Long = 3287089        ' RGB(49, 40, 50) escuro
Private Const LARGURA_MINIMA As Double = 8

Private Function CorCabecalho() As Long
    CorCabecalho = RGB(45, 91, 208)               ' #2d5bd0
End Function

'==============================================================================
' Macro principal
'==============================================================================
Public Sub FormatarDiagnosticoBancos()
    Dim ws As Worksheet
    Dim linhaCab As Long, ultLin As Long, ultCol As Long
    Dim colSituacao As Long, colTotal As Long, colUrl As Long
    Dim colObs As Long, colCmid As Long, colMateria As Long

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    On Error GoTo Falhou

    ' O Excel em inglês abre o CSV de ponto e vírgula tudo na coluna A.
    SepararColunasSePreciso ws

    linhaCab = AchaLinhaCabecalho(ws)
    If linhaCab = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Não encontrei o cabeçalho do diagnóstico (a coluna 'Matéria')." & vbCrLf & vbCrLf & _
               "Abra o CSV gerado pelo botão 'Baixar CSV' do script e rode a macro com ele na frente.", _
               vbExclamation, "Formatar diagnóstico"
        Exit Sub
    End If

    ' Rodando de novo: joga fora o título antigo e a linha de totais antiga,
    ' para remontar tudo em cima dos dados puros.
    If linhaCab > 1 Then
        ws.Rows("1:" & linhaCab - 1).Delete
        linhaCab = 1
    End If
    RemoveLinhaDeTotais ws

    ultCol = ws.Cells(linhaCab, ws.Columns.Count).End(xlToLeft).Column
    ultLin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultLin <= linhaCab Then
        Application.ScreenUpdating = True
        MsgBox "A planilha tem cabeçalho mas nenhuma sala listada.", vbExclamation, "Formatar diagnóstico"
        Exit Sub
    End If

    colMateria = ColunaPorNome(ws, linhaCab, ultCol, "Matéria")
    colCmid = ColunaPorNome(ws, linhaCab, ultCol, "cmid")
    colTotal = ColunaPorNome(ws, linhaCab, ultCol, "Total")
    colSituacao = ColunaPorNome(ws, linhaCab, ultCol, "Situação")
    colUrl = ColunaPorNome(ws, linhaCab, ultCol, "URL")
    colObs = ColunaPorNome(ws, linhaCab, ultCol, "Observações")

    LimpaFormatacaoAntiga ws
    FormataCorpo ws, linhaCab, ultLin, ultCol
    FormataColunas ws, linhaCab, ultLin, ultCol, colCmid, colTotal, colSituacao, colUrl, colObs
    AjustaAlturas ws, linhaCab, ultLin
    PintaPorSituacao ws, linhaCab, ultLin, ultCol, colSituacao
    TransformaUrlsEmLinks ws, linhaCab, ultLin, colUrl
    Dim linhaTotais As Long
    linhaTotais = EscreveLinhaDeTotais(ws, linhaCab, ultLin, ultCol, colCmid, colTotal)
    EscreveTitulo ws, linhaCab, ultLin, ultCol, colSituacao
    ' O título ocupou 3 linhas: tudo desceu.
    linhaCab = linhaCab + 3
    ultLin = ultLin + 3
    linhaTotais = linhaTotais + 3

    FormataCabecalho ws, linhaCab, ultCol
    ws.Range(ws.Cells(linhaCab, 1), ws.Cells(ultLin, ultCol)).AutoFilter
    CongelaPaineis ws, linhaCab, colMateria
    PreparaImpressao ws, linhaCab, linhaTotais, ultCol

    ws.Cells(linhaCab, 1).Select
    Application.ScreenUpdating = True

    OfereceSalvarComoXlsx ws
    Exit Sub

Falhou:
    Application.ScreenUpdating = True
    MsgBox "Não consegui formatar: " & Err.Description, vbCritical, "Formatar diagnóstico"
End Sub

'==============================================================================
' Preparo
'==============================================================================

' Quando o separador do Excel não bate com o do arquivo, a linha inteira cai
' na coluna A. Aqui a gente reparte na mão.
Private Sub SepararColunasSePreciso(ws As Worksheet)
    Dim texto As String
    Dim porPontoEVirgula As Boolean

    If Len(CStr(ws.Range("B1").Value)) > 0 Then Exit Sub
    texto = CStr(ws.Range("A1").Value)
    If InStr(texto, ";") = 0 And InStr(texto, ",") = 0 Then Exit Sub

    porPontoEVirgula = (InStr(texto, ";") > 0)
    ws.Columns(1).TextToColumns _
        Destination:=ws.Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=porPontoEVirgula, Comma:=Not porPontoEVirgula, _
        Space:=False, Other:=False
End Sub

' Procura a linha do cabeçalho nas primeiras linhas. Compara pelo começo do
' texto para não depender do acento em "Matéria".
Private Function AchaLinhaCabecalho(ws As Worksheet) As Long
    Dim i As Long
    For i = 1 To 12
        If Left$(CStr(ws.Cells(i, 1).Value), 3) = "Mat" Then
            AchaLinhaCabecalho = i
            Exit Function
        End If
    Next i
    AchaLinhaCabecalho = 0
End Function

Private Function ColunaPorNome(ws As Worksheet, linhaCab As Long, ultCol As Long, nome As String) As Long
    Dim c As Long
    For c = 1 To ultCol
        If StrComp(Trim$(CStr(ws.Cells(linhaCab, c).Value)), nome, vbTextCompare) = 0 Then
            ColunaPorNome = c
            Exit Function
        End If
    Next c
    ColunaPorNome = 0
End Function

Private Sub RemoveLinhaDeTotais(ws As Worksheet)
    Dim ultLin As Long
    ultLin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultLin < 2 Then Exit Sub
    If Left$(CStr(ws.Cells(ultLin, 1).Value), 5) = "TOTAL" Then ws.Rows(ultLin).Delete
End Sub

Private Sub LimpaFormatacaoAntiga(ws As Worksheet)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.FormatConditions.Delete
    ws.Cells.Interior.Pattern = xlNone
    ws.Cells.Borders.LineStyle = xlNone
    ActiveWindow.FreezePanes = False
End Sub

'==============================================================================
' Aparência
'==============================================================================

Private Sub FormataCorpo(ws As Worksheet, linhaCab As Long, ultLin As Long, ultCol As Long)
    Dim dados As Range
    Set dados = ws.Range(ws.Cells(linhaCab, 1), ws.Cells(ultLin, ultCol))

    With dados
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = RGB(32, 38, 45)
        .VerticalAlignment = xlTop
        .WrapText = False
    End With

    ' Linhas finas cinza: separam sem poluir.
    With dados.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(226, 230, 236)
        .Weight = xlThin
    End With
    With dados.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = RGB(238, 241, 245)
        .Weight = xlThin
    End With

End Sub

' Altura pelo conteúdo (a coluna de observações quebra linha), mas sem deixar
' nenhuma linha mais apertada que o cabeçalho.
Private Sub AjustaAlturas(ws As Worksheet, linhaCab As Long, ultLin As Long)
    Dim i As Long
    ws.Rows(linhaCab + 1 & ":" & ultLin).AutoFit
    For i = linhaCab + 1 To ultLin
        If ws.Rows(i).RowHeight < 18 Then ws.Rows(i).RowHeight = 18
    Next i
End Sub

Private Sub FormataCabecalho(ws As Worksheet, linhaCab As Long, ultCol As Long)
    With ws.Range(ws.Cells(linhaCab, 1), ws.Cells(linhaCab, ultCol))
        .Interior.Color = CorCabecalho()
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(28, 60, 140)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(linhaCab).RowHeight = 30
End Sub

Private Sub FormataColunas(ws As Worksheet, linhaCab As Long, ultLin As Long, ultCol As Long, _
                           colCmid As Long, colTotal As Long, colSituacao As Long, _
                           colUrl As Long, colObs As Long)
    Dim c As Long
    Dim corpo As Range

    For c = 1 To ultCol
        Set corpo = ws.Range(ws.Cells(linhaCab + 1, c), ws.Cells(ultLin, c))

        If colTotal > 0 And colCmid > 0 And c > colCmid And c <= colTotal Then
            ' Contagens de questões e o total: número, centralizado.
            corpo.HorizontalAlignment = xlCenter
            corpo.NumberFormat = "0"
            ws.Columns(c).ColumnWidth = 9
        ElseIf c = colSituacao Then
            corpo.HorizontalAlignment = xlCenter
            corpo.Font.Bold = True
            ws.Columns(c).ColumnWidth = 14
        ElseIf c = colUrl Then
            corpo.HorizontalAlignment = xlCenter
            ws.Columns(c).ColumnWidth = 10
        ElseIf c = colObs Then
            corpo.WrapText = True
            ws.Columns(c).ColumnWidth = 42
        Else
            corpo.HorizontalAlignment = xlLeft
            ws.Columns(c).AutoFit
            If ws.Columns(c).ColumnWidth > 40 Then ws.Columns(c).ColumnWidth = 40
            If ws.Columns(c).ColumnWidth < LARGURA_MINIMA Then ws.Columns(c).ColumnWidth = LARGURA_MINIMA
        End If
    Next c

    ' A primeira coluna é a que a pessoa lê: um pouco maior e em destaque.
    ws.Columns(1).ColumnWidth = 38
    ws.Range(ws.Cells(linhaCab + 1, 1), ws.Cells(ultLin, 1)).Font.Bold = True
End Sub

' Uma regra por situação, aplicada à linha inteira. Fica tudo como formatação
' condicional (e não pintado na mão) para continuar valendo se você editar a
' coluna "Situação" depois.
Private Sub PintaPorSituacao(ws As Worksheet, linhaCab As Long, ultLin As Long, _
                             ultCol As Long, colSituacao As Long)
    Dim dados As Range
    Dim letra As String

    If colSituacao = 0 Then Exit Sub

    Set dados = ws.Range(ws.Cells(linhaCab + 1, 1), ws.Cells(ultLin, ultCol))
    letra = LetraColuna(ws, colSituacao)

    AdicionaRegra dados, letra, linhaCab + 1, "ok", RGB(236, 248, 240), RGB(23, 114, 69)
    AdicionaRegra dados, letra, linhaCab + 1, "incompleto", RGB(255, 247, 230), RGB(138, 97, 0)
    AdicionaRegra dados, letra, linhaCab + 1, "vazio", RGB(253, 237, 235), RGB(179, 38, 30)
    AdicionaRegra dados, letra, linhaCab + 1, "sem banco", RGB(250, 228, 224), RGB(140, 30, 24)
    AdicionaRegra dados, letra, linhaCab + 1, "erro", RGB(240, 242, 244), RGB(90, 102, 115)
End Sub

Private Sub AdicionaRegra(dados As Range, letra As String, primeiraLinha As Long, _
                          situacao As String, corFundo As Long, corTexto As Long)
    Dim regra As FormatCondition
    Set regra = dados.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & letra & primeiraLinha & "=""" & situacao & """")
    regra.Interior.Color = corFundo
    regra.Font.Color = corTexto
    regra.StopIfTrue = False
End Sub

' A URL inteira ocuparia meia tela: vira um "abrir" clicável.
Private Sub TransformaUrlsEmLinks(ws As Worksheet, linhaCab As Long, ultLin As Long, colUrl As Long)
    Dim i As Long
    Dim endereco As String

    If colUrl = 0 Then Exit Sub

    For i = linhaCab + 1 To ultLin
        endereco = Trim$(CStr(ws.Cells(i, colUrl).Value))
        ' Só converte o que ainda é URL: assim dá para rodar a macro de novo.
        If Left$(LCase$(endereco), 4) = "http" Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, colUrl), Address:=endereco, _
                              TextToDisplay:="abrir", ScreenTip:=endereco
        End If
    Next i
End Sub

Private Function EscreveLinhaDeTotais(ws As Worksheet, linhaCab As Long, ultLin As Long, _
                                      ultCol As Long, colCmid As Long, colTotal As Long) As Long
    Dim linha As Long, c As Long
    Dim letra As String

    linha = ultLin + 1
    ws.Cells(linha, 1).Value = "TOTAL (" & (ultLin - linhaCab) & " salas)"

    If colTotal > 0 And colCmid > 0 Then
        For c = colCmid + 1 To colTotal
            letra = LetraColuna(ws, c)
            ws.Cells(linha, c).Formula = "=SUM(" & letra & linhaCab + 1 & ":" & letra & ultLin & ")"
            ws.Cells(linha, c).HorizontalAlignment = xlCenter
            ws.Cells(linha, c).NumberFormat = "0"
        Next c
    End If

    With ws.Range(ws.Cells(linha, 1), ws.Cells(linha, ultCol))
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"
        .Interior.Color = RGB(238, 241, 246)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = CorCabecalho()
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    ws.Rows(linha).RowHeight = 20

    EscreveLinhaDeTotais = linha
End Function

' Três linhas no topo: nome do relatório, contagem por situação e um respiro.
Private Sub EscreveTitulo(ws As Worksheet, linhaCab As Long, ultLin As Long, _
                          ultCol As Long, colSituacao As Long)
    Dim titulo As String

    ws.Rows(linhaCab & ":" & linhaCab + 2).Insert Shift:=xlDown

    titulo = NomeDoRelatorio(ws)
    With ws.Cells(linhaCab, 1)
        .Value = titulo
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 16
        .Font.Color = COR_TITULO
    End With
    ws.Range(ws.Cells(linhaCab, 1), ws.Cells(linhaCab, ultCol)).Merge
    ws.Rows(linhaCab).RowHeight = 26

    With ws.Cells(linhaCab + 1, 1)
        .Value = ResumoPorSituacao(ws, linhaCab + 4, ultLin + 3, colSituacao)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = RGB(90, 102, 115)
    End With
    ws.Range(ws.Cells(linhaCab + 1, 1), ws.Cells(linhaCab + 1, ultCol)).Merge
    ws.Rows(linhaCab + 1).RowHeight = 18
    ws.Rows(linhaCab + 2).RowHeight = 8
End Sub

' O nome do arquivo já traz categoria e data ("diagnostico bancos - X - Y.csv").
Private Function NomeDoRelatorio(ws As Worksheet) As String
    Dim nome As String
    Dim ponto As Long

    nome = ws.Parent.Name
    ponto = InStrRev(nome, ".")
    If ponto > 1 Then nome = Left$(nome, ponto - 1)
    nome = Replace(nome, "diagnostico bancos", "Diagnóstico dos bancos")
    If Len(Trim$(nome)) = 0 Then nome = "Diagnóstico dos bancos"
    NomeDoRelatorio = nome
End Function

Private Function ResumoPorSituacao(ws As Worksheet, primeira As Long, ultima As Long, _
                                   colSituacao As Long) As String
    Dim i As Long
    Dim situacao As String
    Dim ok As Long, incompleto As Long, vazio As Long, semBanco As Long, erro As Long
    Dim partes As String

    If colSituacao = 0 Then
        ResumoPorSituacao = (ultima - primeira + 1) & " sala(s)"
        Exit Function
    End If

    For i = primeira To ultima
        situacao = LCase$(Trim$(CStr(ws.Cells(i, colSituacao).Value)))
        Select Case situacao
            Case "ok": ok = ok + 1
            Case "incompleto": incompleto = incompleto + 1
            Case "vazio": vazio = vazio + 1
            Case "sem banco": semBanco = semBanco + 1
            Case Else: erro = erro + 1
        End Select
    Next i

    partes = (ultima - primeira + 1) & " salas"
    partes = partes & "   ·   " & ok & " completos"
    partes = partes & "   ·   " & incompleto & " incompletos"
    partes = partes & "   ·   " & vazio & " vazios"
    If semBanco > 0 Then partes = partes & "   ·   " & semBanco & " sem banco"
    If erro > 0 Then partes = partes & "   ·   " & erro & " não lidos"
    ResumoPorSituacao = partes
End Function

'==============================================================================
' Janela e impressão
'==============================================================================

Private Sub CongelaPaineis(ws As Worksheet, linhaCab As Long, colMateria As Long)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ' Congela o cabeçalho e a coluna da matéria: rolando para o lado, você
    ' continua sabendo de que sala é cada linha.
    ws.Cells(linhaCab + 1, IIf(colMateria = 1, 2, 1)).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub PreparaImpressao(ws As Worksheet, linhaCab As Long, linhaTotais As Long, ultCol As Long)
    On Error Resume Next   ' sem impressora instalada, o PageSetup reclama
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(linhaTotais, ultCol)).Address
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & linhaCab & ":$" & linhaCab
        .CenterFooter = "&P / &N"
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
    End With
    On Error GoTo 0
End Sub

' CSV não guarda formatação: se salvar por cima, tudo isso se perde.
Private Sub OfereceSalvarComoXlsx(ws As Worksheet)
    Dim wb As Workbook
    Dim destino As String, pasta As String, nome As String
    Dim ponto As Long

    Set wb = ws.Parent
    If LCase$(Right$(wb.Name, 4)) <> ".csv" Then Exit Sub

    If MsgBox("Formatação pronta." & vbCrLf & vbCrLf & _
              "O arquivo aberto é um CSV, e CSV não guarda formatação: se salvar assim, " & _
              "as cores e larguras somem." & vbCrLf & vbCrLf & _
              "Salvar como .xlsx na mesma pasta?", _
              vbYesNo + vbQuestion, "Formatar diagnóstico") <> vbYes Then Exit Sub

    nome = wb.Name
    ponto = InStrRev(nome, ".")
    If ponto > 1 Then nome = Left$(nome, ponto - 1)

    pasta = wb.Path
    If Len(pasta) = 0 Then pasta = Application.DefaultFilePath

    destino = pasta & Application.PathSeparator & nome & ".xlsx"
    ' DisplayAlerts fica ligado de propósito: se o arquivo já existir, quem
    ' decide sobrescrever é você, não a macro.
    wb.SaveAs Filename:=destino, FileFormat:=xlOpenXMLWorkbook
End Sub

'==============================================================================
' Utilidades
'==============================================================================

Private Function LetraColuna(ws As Worksheet, col As Long) As String
    LetraColuna = Split(ws.Cells(1, col).Address(True, False), "$")(0)
End Function
