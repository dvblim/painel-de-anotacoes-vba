VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepositorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Única parte do projeto que lê e grava nas tabelas da planilha.
' O painel nunca mexe direto nas células: sempre passa por aqui.

Private Const COL_ID As Long = 1
Private Const COL_DATA As Long = 2
Private Const COL_HORA As Long = 3
Private Const COL_TITULO As Long = 4
Private Const COL_CONTEUDO As Long = 5
Private Const COL_URGENTE As Long = 6
Private Const COL_REGISTRO As Long = 7
Private Const COL_ALTERADO As Long = 8
Private Const COL_LEMBRETE As Long = 9
Private Const COL_CONCLUIDO As Long = 10
Private Const COL_FINALIZADA As Long = 11
Private Const COL_FINALIZADA_EM As Long = 12

' ---------- acesso às tabelas ----------

Private Sub Class_Initialize()
    GarantirColunas
End Sub

' Bases criadas por versões anteriores não têm as colunas do lembrete
Private Sub GarantirColunas()
    Dim lo As ListObject
    Set lo = Tabela
    If lo.ListColumns.Count < COL_LEMBRETE Then
        lo.ListColumns.Add.Name = "Lembrete"
        lo.ListColumns(COL_LEMBRETE).Range.ColumnWidth = 18
    End If
    If lo.ListColumns.Count < COL_CONCLUIDO Then
        lo.ListColumns.Add.Name = "Lembrete concluído"
        lo.ListColumns(COL_CONCLUIDO).Range.ColumnWidth = 20
    End If
    If lo.ListColumns.Count < COL_FINALIZADA Then
        lo.ListColumns.Add.Name = "Finalizada"
        lo.ListColumns(COL_FINALIZADA).Range.ColumnWidth = 12
    End If
    If lo.ListColumns.Count < COL_FINALIZADA_EM Then
        lo.ListColumns.Add.Name = "Finalizada em"
        lo.ListColumns(COL_FINALIZADA_EM).Range.ColumnWidth = 18
    End If
End Sub

Private Function Tabela() As ListObject
    Set Tabela = ThisWorkbook.Worksheets(ABA_BASE).ListObjects(TB_NOTAS)
End Function

Private Function Dados() As Variant
    Dim lo As ListObject
    Set lo = Tabela
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dados = lo.DataBodyRange.Value
End Function

Private Function TextoDe(ByVal x As Variant) As String
    If IsError(x) Then Exit Function
    If IsNull(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    TextoDe = CStr(x)
End Function

Private Function EhData(ByVal x As Variant) As Boolean
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then Exit Function
    If Len(Trim$(CStr(x))) = 0 Then Exit Function
    EhData = IsDate(x) Or IsNumeric(x)
End Function

Private Function DiaDe(ByVal x As Variant) As Long
    DiaDe = CLng(Fix(CDbl(CDate(x))))
End Function

Private Function HoraTexto(ByVal x As Variant) As String
    If Len(Trim$(TextoDe(x))) = 0 Then Exit Function
    If IsDate(x) Or IsNumeric(x) Then
        HoraTexto = Format$(CDate(x), "hh:nn")
    Else
        HoraTexto = CStr(x)
    End If
End Function

Private Function LinhaValida(ByRef v As Variant, ByVal i As Long) As Boolean
    If Len(Trim$(TextoDe(v(i, COL_ID)))) = 0 Then Exit Function
    If Not IsNumeric(v(i, COL_ID)) Then Exit Function
    If Not EhData(v(i, COL_DATA)) Then Exit Function
    LinhaValida = True
End Function

Private Function NotaDaLinha(ByRef v As Variant, ByVal i As Long) As clsAnotacao
    Dim n As clsAnotacao
    Set n = New clsAnotacao
    n.ID = CLng(v(i, COL_ID))
    n.Data = CDate(DiaDe(v(i, COL_DATA)))
    n.Hora = HoraTexto(v(i, COL_HORA))
    n.Titulo = TextoDe(v(i, COL_TITULO))
    n.Conteudo = Replace(Replace(TextoDe(v(i, COL_CONTEUDO)), vbCrLf, vbLf), vbLf, vbCrLf)
    n.Urgente = (UCase$(Trim$(TextoDe(v(i, COL_URGENTE)))) = "SIM")
    If EhData(v(i, COL_LEMBRETE)) Then n.Lembrete = CDate(v(i, COL_LEMBRETE))
    n.LembreteConcluido = (UCase$(Trim$(TextoDe(v(i, COL_CONCLUIDO)))) = "SIM")
    n.Finalizada = (UCase$(Trim$(TextoDe(v(i, COL_FINALIZADA)))) = "SIM")
    If EhData(v(i, COL_FINALIZADA_EM)) Then n.FinalizadaEm = CDate(v(i, COL_FINALIZADA_EM))
    Set NotaDaLinha = n
End Function

' Evita que um texto iniciado por = + - @ seja interpretado como fórmula
Private Function TextoSeguro(ByVal s As String) As String
    If Len(s) > 0 Then
        If InStr(1, "=+-@", Left$(s, 1)) > 0 Then s = "'" & s
    End If
    TextoSeguro = s
End Function

Private Function IndiceDoID(ByVal idNota As Long) As Long
    Dim v As Variant, i As Long
    v = Dados
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        If LinhaValida(v, i) Then
            If CLng(v(i, COL_ID)) = idNota Then
                IndiceDoID = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ProximoID() As Long
    Dim v As Variant, i As Long, maior As Long
    v = Dados
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If LinhaValida(v, i) Then
                If CLng(v(i, COL_ID)) > maior Then maior = CLng(v(i, COL_ID))
            End If
        Next i
    End If
    ProximoID = maior + 1
End Function

' ---------- consultas ----------

Public Function ListarPorData(ByVal dia As Date) As Collection
    Dim col As Collection, v As Variant, i As Long
    Set col = New Collection
    v = Dados
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If LinhaValida(v, i) Then
                If DiaDe(v(i, COL_DATA)) = CLng(dia) Then col.Add NotaDaLinha(v, i)
            End If
        Next i
    End If
    Set ListarPorData = col
End Function

Private Sub InserirPorData(ByVal col As Collection, ByVal n As clsAnotacao)
    Dim j As Long, existente As clsAnotacao
    For j = 1 To col.Count
        Set existente = col(j)
        If n.Data < existente.Data Then
            col.Add n, , j
            Exit Sub
        End If
    Next j
    col.Add n
End Sub

' Anotações com lembrete ativo (primeiro) e as marcadas como urgentes (depois),
' cada grupo da data mais antiga para a mais nova; as finalizadas ficam de fora
Public Function ListarAvisos() As Collection
    Dim ativos As Collection, urgentes As Collection, v As Variant, i As Long
    Dim n As clsAnotacao
    Set ativos = New Collection
    Set urgentes = New Collection
    v = Dados
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If LinhaValida(v, i) Then
                Set n = NotaDaLinha(v, i)
                If n.LembreteAtivo Then
                    InserirPorData ativos, n
                ElseIf n.Urgente And Not n.Finalizada Then
                    InserirPorData urgentes, n
                End If
            End If
        Next i
    End If
    For i = 1 To urgentes.Count
        ativos.Add urgentes(i)
    Next i
    Set ListarAvisos = ativos
End Function

Public Function Obter(ByVal idNota As Long) As clsAnotacao
    Dim v As Variant, i As Long
    i = IndiceDoID(idNota)
    If i = 0 Then Exit Function
    v = Dados
    Set Obter = NotaDaLinha(v, i)
End Function

' Dicionário dia-do-mês -> soma de 1 (tem anotação), 2 (tem urgente em aberto),
' 4 (tem lembrete ativo) e 8 (tem anotação ainda não finalizada)
Public Function MarcasDoMes(ByVal ano As Long, ByVal mes As Long) As Object
    Dim dic As Object, v As Variant, i As Long
    Dim n As clsAnotacao, chave As Long, marcas As Long
    Set dic = CreateObject("Scripting.Dictionary")
    v = Dados
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If LinhaValida(v, i) Then
                Set n = NotaDaLinha(v, i)
                If Year(n.Data) = ano And Month(n.Data) = mes Then
                    chave = CLng(Day(n.Data))
                    marcas = 1
                    If n.Urgente And Not n.Finalizada Then marcas = marcas Or 2
                    If Not n.Finalizada Then marcas = marcas Or 8
                    If n.LembreteAtivo Then marcas = marcas Or 4
                    If dic.Exists(chave) Then marcas = marcas Or CLng(dic(chave))
                    dic(chave) = marcas
                End If
            End If
        Next i
    End If
    Set MarcasDoMes = dic
End Function

' Dicionário data (número de série) -> descrição do feriado
Public Function Feriados() As Object
    Dim dic As Object, lo As ListObject, v As Variant, i As Long
    Set dic = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(ABA_FERIADOS).ListObjects(TB_FERIADOS)
    On Error GoTo 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            v = lo.DataBodyRange.Value
            For i = 1 To UBound(v, 1)
                If EhData(v(i, 1)) Then dic(DiaDe(v(i, 1))) = TextoDe(v(i, 2))
            Next i
        End If
    End If
    Set Feriados = dic
End Function

' ---------- gravação ----------

' Grava uma anotação nova (ID = 0) ou atualiza uma existente
Public Sub Salvar(ByVal n As clsAnotacao)
    Dim lo As ListObject, r As Range, i As Long, nova As Boolean

    Set lo = Tabela
    If n.ID <> 0 Then i = IndiceDoID(n.ID)

    If i > 0 Then
        Set r = lo.ListRows(i).Range
    Else
        nova = True
        n.ID = ProximoID
        If lo.ListRows.Count = 1 Then
            ' reaproveita a linha em branco de uma tabela vazia
            If Len(Trim$(TextoDe(lo.DataBodyRange.Cells(1, COL_ID).Value))) = 0 Then
                Set r = lo.ListRows(1).Range
            End If
        End If
        If r Is Nothing Then Set r = lo.ListRows.Add.Range
    End If

    If nova Then
        r.Cells(1, COL_ID).Value = n.ID
        r.Cells(1, COL_DATA).NumberFormat = "dd/mm/yyyy"
        r.Cells(1, COL_DATA).Value = n.Data
        r.Cells(1, COL_HORA).NumberFormat = "hh:mm"
        If Len(n.Hora) > 0 Then r.Cells(1, COL_HORA).Value = TimeValue(n.Hora)
        r.Cells(1, COL_REGISTRO).NumberFormat = "dd/mm/yyyy hh:mm"
        r.Cells(1, COL_REGISTRO).Value = Now
    Else
        r.Cells(1, COL_ALTERADO).NumberFormat = "dd/mm/yyyy hh:mm"
        r.Cells(1, COL_ALTERADO).Value = Now
    End If

    r.Cells(1, COL_TITULO).Value = TextoSeguro(n.Titulo)
    r.Cells(1, COL_CONTEUDO).Value = TextoSeguro(Replace(n.Conteudo, vbCrLf, vbLf))
    r.Cells(1, COL_URGENTE).Value = IIf(n.Urgente, "SIM", "NÃO")

    If n.Lembrete > 0 Then
        r.Cells(1, COL_LEMBRETE).NumberFormat = "dd/mm/yyyy hh:mm"
        r.Cells(1, COL_LEMBRETE).Value = n.Lembrete
        r.Cells(1, COL_CONCLUIDO).Value = IIf(n.LembreteConcluido, "SIM", "NÃO")
    Else
        r.Cells(1, COL_LEMBRETE).ClearContents
        r.Cells(1, COL_CONCLUIDO).ClearContents
    End If

    If n.Finalizada Then
        r.Cells(1, COL_FINALIZADA).Value = "SIM"
        r.Cells(1, COL_FINALIZADA_EM).NumberFormat = "dd/mm/yyyy hh:mm"
        If n.FinalizadaEm > 0 Then r.Cells(1, COL_FINALIZADA_EM).Value = n.FinalizadaEm
    Else
        r.Cells(1, COL_FINALIZADA).Value = "NÃO"
        r.Cells(1, COL_FINALIZADA_EM).ClearContents
    End If
End Sub

' Dá baixa na anotação ou reabre, se já estiver finalizada
Public Sub AlternarFinalizada(ByVal idNota As Long)
    Dim n As clsAnotacao
    Set n = Obter(idNota)
    If n Is Nothing Then Exit Sub
    n.Finalizada = Not n.Finalizada
    If n.Finalizada Then n.FinalizadaEm = Now Else n.FinalizadaEm = 0
    Salvar n
End Sub

Public Sub AlternarUrgente(ByVal idNota As Long)
    Dim i As Long, c As Range
    i = IndiceDoID(idNota)
    If i = 0 Then Exit Sub
    Set c = Tabela.ListRows(i).Range.Cells(1, COL_URGENTE)
    If UCase$(Trim$(TextoDe(c.Value))) = "SIM" Then
        c.Value = "NÃO"
    Else
        c.Value = "SIM"
    End If
End Sub

Public Sub Excluir(ByVal idNota As Long)
    Dim i As Long
    i = IndiceDoID(idNota)
    If i = 0 Then Exit Sub
    Tabela.ListRows(i).Delete
End Sub

' Salva o arquivo para que a anotação não se perca se o Excel for fechado
Public Sub Persistir()
    On Error Resume Next
    ThisWorkbook.Save
End Sub
