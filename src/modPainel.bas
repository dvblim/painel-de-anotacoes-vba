Attribute VB_Name = "modPainel"
Option Explicit

' Nomes da aba e das tabelas usadas como base de dados
Public Const ABA_BASE As String = "Base"
Public Const TB_NOTAS As String = "tbAnotacoes"
Public Const ABA_FERIADOS As String = "Feriados"
Public Const TB_FERIADOS As String = "tbFeriados"
Public Const ABA_PAINEL As String = "Painel"

' Executa sozinho ao abrir a pasta de trabalho
Public Sub Auto_Open()
    On Error Resume Next
    GarantirBotao
    AvisarLembretes
End Sub

' Abre o painel (também pode ser chamado por Alt+F8)
Public Sub AbrirPainel()
    On Error GoTo Falha
    frmPainel.Show
    Exit Sub
Falha:
    MsgBox "Não foi possível abrir o painel." & vbCrLf & vbCrLf & _
           "Detalhe: " & Err.Description, vbExclamation, "Painel de Anotações"
End Sub

' Cria o botão ABRIR PAINEL na aba Painel, caso ainda não exista
Public Sub GarantirBotao()
    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Worksheets(ABA_PAINEL)

    On Error Resume Next
    Set shp = ws.Shapes("btnAbrirPainel")
    On Error GoTo 0
    If Not shp Is Nothing Then Exit Sub

    Set shp = ws.Shapes.AddShape(1, ws.Range("B4").Left, ws.Range("B4").Top, 190, 38)
    With shp
        .Name = "btnAbrirPainel"
        .Fill.ForeColor.RGB = RGB(43, 52, 69)
        .Line.ForeColor.RGB = RGB(242, 194, 48)
        .Line.Weight = 1.5
        .TextFrame2.TextRange.Text = "ABRIR PAINEL"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame2.VerticalAnchor = 3
        .OnAction = "AbrirPainel"
    End With
End Sub

' Ao abrir o arquivo, lista os lembretes que já estão valendo
Public Sub AvisarLembretes()
    Dim repo As clsRepositorio, avisos As Collection, n As clsAnotacao
    Dim i As Long, texto As String, total As Long

    Set repo = New clsRepositorio
    Set avisos = repo.ListarAvisos
    For i = 1 To avisos.Count
        Set n = avisos(i)
        If n.LembreteAtivo Then
            total = total + 1
            If total <= 10 Then
                texto = texto & "!  " & Format$(n.Data, "dd\/mm\/yyyy") & " · " & n.Titulo & vbCrLf
            End If
        End If
    Next i
    If total = 0 Then Exit Sub
    If total > 10 Then texto = texto & "... e mais " & (total - 10) & "." & vbCrLf

    MsgBox "Lembretes ativos:" & vbCrLf & vbCrLf & texto & vbCrLf & _
           "Abra o painel para ver, editar ou concluir.", vbExclamation, "Painel de Anotações"
End Sub

' ---------- nomes de dias ----------

Public Function NomeDia(ByVal dia As Date) As String
    NomeDia = Array("domingo", "segunda-feira", "terça-feira", "quarta-feira", _
                    "quinta-feira", "sexta-feira", "sábado")(Weekday(dia, vbSunday) - 1)
End Function

Public Function DiaAbrev(ByVal dia As Date) As String
    DiaAbrev = Array("dom", "seg", "ter", "qua", "qui", "sex", "sáb")(Weekday(dia, vbSunday) - 1)
End Function

' ---------- caixas de listagem do lembrete (dia e horário) ----------

Private Function RotuloDia(ByVal diaNota As Date, ByVal prazo As Long) As String
    Dim d As Date, s As String
    d = diaNota - prazo
    s = Format$(d, "dd\/mm") & " " & DiaAbrev(d) & " · "
    If prazo = 0 Then
        s = s & "no dia"
    ElseIf prazo = 1 Then
        s = s & "1 dia antes"
    ElseIf prazo > 1 Then
        s = s & prazo & " dias antes"
    Else
        s = s & "depois da data"
    End If
    RotuloDia = s
End Function

' Lista "sem lembrete", o dia da anotação e os dias anteriores.
' A segunda coluna (oculta) guarda quantos dias antes.
Public Sub PreencherDias(ByVal cbo As Object, ByVal diaNota As Date)
    Dim prazos As Variant, i As Long
    prazos = Array(0, 1, 2, 3, 4, 5, 6, 7, 15, 30)
    cbo.Clear
    cbo.AddItem "sem lembrete"
    cbo.List(0, 1) = "x"
    For i = LBound(prazos) To UBound(prazos)
        cbo.AddItem RotuloDia(diaNota, CLng(prazos(i)))
        cbo.List(cbo.ListCount - 1, 1) = CStr(prazos(i))
    Next i
    cbo.ListIndex = 0
End Sub

' Horários de 15 em 15 minutos; começa em 09:00
Public Sub PreencherHoras(ByVal cbo As Object)
    Dim m As Long
    cbo.Clear
    For m = 0 To 1425 Step 15
        cbo.AddItem Format$(TimeSerial(0, m, 0), "hh:nn")
    Next m
    cbo.ListIndex = 36
End Sub

' Dia e hora escolhidos nas caixas; devolve 0 quando está em "sem lembrete"
Public Function LembreteDosCombos(ByVal cboDia As Object, ByVal cboHora As Object, ByVal diaNota As Date) As Date
    Dim prazo As Long, hora As String
    If cboDia.ListIndex <= 0 Then Exit Function
    prazo = CLng(cboDia.List(cboDia.ListIndex, 1))
    If cboHora.ListIndex >= 0 Then hora = CStr(cboHora.List(cboHora.ListIndex, 0)) Else hora = "09:00"
    LembreteDosCombos = CDate(CLng(diaNota) - prazo) + TimeValue(hora)
End Function

' Preenche as caixas e posiciona no lembrete gravado (se houver)
Public Sub SelecionarCombos(ByVal cboDia As Object, ByVal cboHora As Object, ByVal diaNota As Date, ByVal lembrete As Date)
    Dim prazo As Long, hora As String, i As Long, achou As Boolean

    PreencherDias cboDia, diaNota
    PreencherHoras cboHora
    If lembrete = 0 Then Exit Sub

    prazo = CLng(diaNota) - CLng(Fix(CDbl(lembrete)))
    For i = 1 To cboDia.ListCount - 1
        If CStr(cboDia.List(i, 1)) = CStr(prazo) Then
            cboDia.ListIndex = i
            achou = True
            Exit For
        End If
    Next i
    If Not achou Then
        cboDia.AddItem RotuloDia(diaNota, prazo)
        cboDia.List(cboDia.ListCount - 1, 1) = CStr(prazo)
        cboDia.ListIndex = cboDia.ListCount - 1
    End If

    hora = Format$(lembrete, "hh:nn")
    achou = False
    For i = 0 To cboHora.ListCount - 1
        If CStr(cboHora.List(i, 0)) = hora Then
            cboHora.ListIndex = i
            achou = True
            Exit For
        End If
    Next i
    If Not achou Then
        cboHora.AddItem hora
        cboHora.ListIndex = cboHora.ListCount - 1
    End If
End Sub
