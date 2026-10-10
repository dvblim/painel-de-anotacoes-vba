VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPainel
   Caption         =   "Painel de Anotações"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15900
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPainel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Painel de anotações por data.
' Todos os controles são criados por código em UserForm_Initialize,
' por isso o formulário não depende de arquivo .frx.

Private mRepo As clsRepositorio
Private mDias As Collection          ' 42 x clsDiaCalendario
Private mCartoes As Collection       ' clsCartaoNota, criados conforme a necessidade
Private mItens As Collection         ' clsItemAviso, idem
Private mAno As Long
Private mMes As Long
Private mDiaSel As Date
Private mEditandoID As Long

Private mCorNavy As Long
Private mCorTexto As Long
Private mCorCinza As Long
Private mCorBorda As Long

Private WithEvents lblAnterior As MSForms.Label
Attribute lblAnterior.VB_VarHelpID = -1
Private WithEvents lblProximo As MSForms.Label
Attribute lblProximo.VB_VarHelpID = -1
Private WithEvents lblHoje As MSForms.Label
Attribute lblHoje.VB_VarHelpID = -1
Private WithEvents btnSalvar As MSForms.CommandButton
Attribute btnSalvar.VB_VarHelpID = -1
Private WithEvents btnCancelar As MSForms.CommandButton
Attribute btnCancelar.VB_VarHelpID = -1

Private lblMes As MSForms.Label
Private lblFeriados As MSForms.Label
Private lblTitNotas As MSForms.Label
Private lblContagem As MSForms.Label
Private lblVazio As MSForms.Label
Private lblTitNova As MSForms.Label
Private fraLista As MSForms.Frame
Private fraAvisos As MSForms.Frame
Private lblSemAvisos As MSForms.Label
Private txtTitulo As MSForms.TextBox
Private txtConteudo As MSForms.TextBox
Private chkUrgente As MSForms.CheckBox
Private cboDia As MSForms.ComboBox
Private cboHora As MSForms.ComboBox

' ============================================================
'  Montagem da tela
' ============================================================

Private Sub UserForm_Initialize()
    mCorNavy = RGB(43, 52, 69)
    mCorTexto = RGB(232, 236, 241)
    mCorCinza = RGB(154, 163, 178)
    mCorBorda = RGB(58, 67, 84)

    Me.Caption = "Painel de Anotações"
    Me.BackColor = RGB(24, 28, 36)
    Me.Width = 930
    Me.Height = 478

    Set mRepo = New clsRepositorio
    Set mDias = New Collection
    Set mCartoes = New Collection
    Set mItens = New Collection

    MontarCalendario
    MontarUrgentes
    MontarLista
    MontarNovaAnotacao

    mDiaSel = Date
    PreencherDias cboDia, mDiaSel
    PreencherHoras cboHora
    mAno = Year(mDiaSel)
    mMes = Month(mDiaSel)

    DesenharCalendario
    CarregarNotas False
    CarregarUrgentes
End Sub

Private Function NovoRotulo(ByVal pai As Object, ByVal nome As String, ByVal texto As String, _
                            ByVal esq As Single, ByVal topo As Single, ByVal larg As Single, ByVal alt As Single, _
                            Optional ByVal tam As Single = 9, Optional ByVal negrito As Boolean = False) As MSForms.Label
    Dim c As MSForms.Label
    Set c = pai.Controls.Add("Forms.Label.1", nome, True)
    With c
        .Caption = texto
        .Left = esq
        .Top = topo
        .Width = larg
        .Height = alt
        .Font.Name = "Segoe UI"
        .Font.Size = tam
        .Font.Bold = negrito
        .ForeColor = mCorTexto
        .BackStyle = fmBackStyleTransparent
    End With
    Set NovoRotulo = c
End Function

Private Function NovoBloco(ByVal nome As String, ByVal esq As Single, ByVal topo As Single, _
                           ByVal larg As Single, ByVal alt As Single, ByVal cor As Long) As MSForms.Label
    Dim c As MSForms.Label
    Set c = NovoRotulo(Me, nome, "", esq, topo, larg, alt)
    c.BackStyle = fmBackStyleOpaque
    c.BackColor = cor
    Set NovoBloco = c
End Function

Private Sub MontarCalendario()
    Const ESQ As Single = 14
    Dim i As Long, lin As Long, col As Long
    Dim iniciais As Variant, casa As clsDiaCalendario, c As MSForms.Label

    NovoRotulo Me, "lblTitCal", "CALENDÁRIO", ESQ, 10, 252, 16, 11, True

    ' barra do mês com filete amarelo
    NovoBloco "blkBarraMes", ESQ, 32, 252, 24, mCorNavy
    NovoBloco "blkFilete", ESQ, 56, 252, 2, RGB(242, 194, 48)

    Set lblAnterior = NovoRotulo(Me, "lblAnterior", "<", ESQ, 36, 30, 16, 10, True)
    lblAnterior.ForeColor = RGB(255, 255, 255)
    lblAnterior.TextAlign = fmTextAlignCenter
    lblAnterior.ControlTipText = "Mês anterior"

    Set lblProximo = NovoRotulo(Me, "lblProximo", ">", ESQ + 222, 36, 30, 16, 10, True)
    lblProximo.ForeColor = RGB(255, 255, 255)
    lblProximo.TextAlign = fmTextAlignCenter
    lblProximo.ControlTipText = "Próximo mês"

    Set lblMes = NovoRotulo(Me, "lblMes", "", ESQ + 30, 36, 192, 16, 10, True)
    lblMes.ForeColor = RGB(255, 255, 255)
    lblMes.TextAlign = fmTextAlignCenter

    ' iniciais dos dias da semana
    iniciais = Array("D", "S", "T", "Q", "Q", "S", "S")
    For col = 0 To 6
        Set c = NovoRotulo(Me, "lblSem" & col, CStr(iniciais(col)), ESQ + col * 36, 64, 34, 12, 8, True)
        c.ForeColor = mCorCinza
        c.TextAlign = fmTextAlignCenter
    Next col

    ' 42 casas (6 semanas)
    For i = 1 To 42
        lin = (i - 1) \ 7
        col = (i - 1) Mod 7
        Set casa = New clsDiaCalendario
        casa.Criar Me, i, ESQ + col * 36, 80 + lin * 25
        mDias.Add casa
    Next i

    ' legenda
    NovoBloco "blkLegFeriado", ESQ, 238, 8, 8, RGB(110, 38, 38)
    Set c = NovoRotulo(Me, "lblLegFeriado", "feriado", ESQ + 12, 235, 60, 12, 8)
    c.ForeColor = mCorCinza
    Set c = NovoBloco("blkLegFds", ESQ + 80, 238, 8, 8, RGB(28, 33, 42))
    c.BorderStyle = fmBorderStyleSingle
    c.BorderColor = mCorBorda
    Set c = NovoRotulo(Me, "lblLegFds", "fim de semana", ESQ + 92, 235, 70, 12, 8)
    c.ForeColor = mCorCinza
    NovoBloco "blkLegNota", ESQ, 252, 5, 9, RGB(200, 206, 216)
    Set c = NovoRotulo(Me, "lblLegNota", "tem anotação", ESQ + 12, 250, 66, 12, 8)
    c.ForeColor = mCorCinza
    NovoBloco "blkLegUrg", ESQ + 80, 252, 5, 9, RGB(239, 83, 80)
    Set c = NovoRotulo(Me, "lblLegUrg", "tem urgente", ESQ + 92, 250, 70, 12, 8)
    c.ForeColor = mCorCinza

    NovoBloco "blkLegLembrete", ESQ + 176, 237, 5, 9, RGB(242, 194, 48)
    Set c = NovoRotulo(Me, "lblLegLembrete", "lembrete", ESQ + 186, 235, 60, 12, 8)
    c.ForeColor = mCorCinza

    NovoBloco "blkLegFinal", ESQ + 176, 252, 5, 9, RGB(63, 185, 80)
    Set c = NovoRotulo(Me, "lblLegFinal", "finalizado", ESQ + 186, 250, 60, 12, 8)
    c.ForeColor = mCorCinza

    Set lblHoje = NovoRotulo(Me, "lblHoje", "ir para hoje", ESQ + 196, 13, 56, 12, 8)
    lblHoje.ForeColor = RGB(125, 170, 240)
    lblHoje.Font.Underline = True
    lblHoje.TextAlign = fmTextAlignRight

    Set lblFeriados = NovoRotulo(Me, "lblFeriados", "", ESQ, 268, 252, 36, 8, True)
    lblFeriados.WordWrap = True
End Sub

Private Sub MontarUrgentes()
    Const ESQ As Single = 14
    NovoRotulo Me, "lblTitUrg", "ALERTAS E URGENTES", ESQ, 310, 252, 16, 11, True

    ' quadro rolável: cada aviso aparece inteiro, quebrando em mais de uma linha se preciso
    Set fraAvisos = Me.Controls.Add("Forms.Frame.1", "fraAvisos", True)
    With fraAvisos
        .Caption = ""
        .Left = ESQ
        .Top = 330
        .Width = 252
        .Height = 110
        .BackColor = RGB(32, 38, 49)
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = mCorBorda
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsNone
    End With

    Set lblSemAvisos = NovoRotulo(fraAvisos, "lblSemAvisos", "Nenhum alerta ou urgente em aberto.", 8, 8, 220, 12, 8)
    lblSemAvisos.ForeColor = mCorCinza
End Sub

Private Sub MontarLista()
    Const ESQ As Single = 284

    Set lblTitNotas = NovoRotulo(Me, "lblTitNotas", "", ESQ, 10, 350, 16, 11, True)
    Set lblContagem = NovoRotulo(Me, "lblContagem", "", ESQ + 354, 13, 266, 12, 8, True)
    lblContagem.ForeColor = mCorCinza
    lblContagem.TextAlign = fmTextAlignRight

    Set fraLista = Me.Controls.Add("Forms.Frame.1", "fraLista", True)
    With fraLista
        .Caption = ""
        .Left = ESQ
        .Top = 32
        .Width = 620
        .Height = 190
        .BackColor = RGB(32, 38, 49)
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = mCorBorda
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With

    Set lblVazio = NovoRotulo(fraLista, "lblVazio", _
        "Nenhuma anotação nesta data. Use o campo abaixo para criar a primeira.", 10, 10, 440, 14, 9)
    lblVazio.ForeColor = mCorCinza
End Sub

Private Sub MontarNovaAnotacao()
    Const ESQ As Single = 284
    Dim c As MSForms.Label

    Set lblTitNova = NovoRotulo(Me, "lblTitNova", "NOVA ANOTAÇÃO", ESQ, 230, 170, 16, 11, True)
    Set c = NovoRotulo(Me, "lblDicaNova", "(data e hora são inseridas automaticamente ao salvar)", _
                       ESQ + 174, 233, 326, 12, 8, True)
    c.ForeColor = mCorCinza

    ' fundo do campo de nova anotação, com contorno âmbar
    Set c = NovoBloco("blkNova", ESQ, 250, 620, 190, RGB(32, 38, 49))
    c.BorderStyle = fmBorderStyleSingle
    c.BorderColor = RGB(120, 100, 40)

    Set c = NovoRotulo(Me, "lblCampoTitulo", "Título", ESQ + 10, 256, 100, 12, 8, True)
    c.ForeColor = mCorCinza

    Set txtTitulo = Me.Controls.Add("Forms.TextBox.1", "txtTitulo", True)
    With txtTitulo
        .Left = ESQ + 10
        .Top = 270
        .Width = 316
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(74, 84, 104)
        .BackColor = RGB(24, 28, 36)
        .ForeColor = mCorTexto
        .MaxLength = 200
    End With

    Set c = NovoRotulo(Me, "lblCampoConteudo", "Conteúdo", ESQ + 10, 296, 100, 12, 8, True)
    c.ForeColor = mCorCinza

    Set txtConteudo = Me.Controls.Add("Forms.TextBox.1", "txtConteudo", True)
    With txtConteudo
        .Left = ESQ + 10
        .Top = 310
        .Width = 600
        .Height = 92
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .MultiLine = True
        .WordWrap = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(74, 84, 104)
        .BackColor = RGB(24, 28, 36)
        .ForeColor = mCorTexto
    End With

    ' lembrete na linha do título: as listas abrem sobre a caixa de conteúdo
    Set c = NovoRotulo(Me, "lblLembrarEm", "Lembrar em", ESQ + 336, 256, 100, 12, 8, True)
    c.ForeColor = RGB(242, 194, 48)
    Set cboDia = NovoCombo("cboDia", ESQ + 336, 190)
    cboDia.ColumnCount = 2
    cboDia.ColumnWidths = "180 pt;0 pt"
    Set c = NovoRotulo(Me, "lblHorario", "Horário", ESQ + 536, 256, 74, 12, 8, True)
    c.ForeColor = RGB(242, 194, 48)
    Set cboHora = NovoCombo("cboHora", ESQ + 536, 74)

    Set chkUrgente = Me.Controls.Add("Forms.CheckBox.1", "chkUrgente", True)
    With chkUrgente
        .Left = ESQ + 10
        .Top = 412
        .Width = 170
        .Height = 18
        .Caption = "Marcar como urgente"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = RGB(32, 38, 49)
        .ForeColor = mCorTexto
        .Value = False
    End With

    Set btnCancelar = Me.Controls.Add("Forms.CommandButton.1", "btnCancelar", True)
    With btnCancelar
        .Left = ESQ + 352
        .Top = 409
        .Width = 110
        .Height = 24
        .Caption = "Cancelar edição"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackColor = RGB(58, 67, 84)
        .ForeColor = mCorTexto
        .Visible = False
    End With

    Set btnSalvar = Me.Controls.Add("Forms.CommandButton.1", "btnSalvar", True)
    With btnSalvar
        .Left = ESQ + 470
        .Top = 409
        .Width = 140
        .Height = 24
        .Caption = "SALVAR ANOTAÇÃO"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = RGB(242, 194, 48)
        .ForeColor = RGB(24, 28, 36)
    End With
End Sub

Private Function NovoCombo(ByVal nome As String, ByVal esq As Single, ByVal larg As Single) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    With c
        .Left = esq
        .Top = 271
        .Width = larg
        .Height = 18
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .BackColor = RGB(24, 28, 36)
        .ForeColor = mCorTexto
        .ListRows = 12
    End With
    Set NovoCombo = c
End Function

' ============================================================
'  Calendário
' ============================================================

Private Function NomeMes(ByVal mes As Long) As String
    NomeMes = Array("JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", _
                    "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")(mes - 1)
End Function

Private Function NomeDiaSemana(ByVal dia As Date) As String
    NomeDiaSemana = Array("domingo", "segunda-feira", "terça-feira", "quarta-feira", _
                          "quinta-feira", "sexta-feira", "sábado")(Weekday(dia, vbSunday) - 1)
End Function

Private Sub DesenharCalendario()
    Dim primeiro As Date, desloc As Long, i As Long, d As Date
    Dim marcas As Object, fer As Object
    Dim casa As clsDiaCalendario, marca As Long, chave As Long
    Dim textoFer As String

    primeiro = DateSerial(mAno, mMes, 1)
    desloc = Weekday(primeiro, vbSunday) - 1
    Set marcas = mRepo.MarcasDoMes(mAno, mMes)
    Set fer = mRepo.Feriados

    lblMes.Caption = NomeMes(mMes) & " " & mAno

    For i = 1 To 42
        Set casa = mDias(i)
        d = primeiro + (i - 1 - desloc)
        If Year(d) = mAno And Month(d) = mMes Then
            chave = CLng(Day(d))
            marca = 0
            If marcas.Exists(chave) Then marca = CLng(marcas(chave))
            casa.Mostrar d, (d = mDiaSel), (d = Date), fer.Exists(CLng(d)), marca
            If fer.Exists(CLng(d)) Then
                If Len(textoFer) > 0 Then textoFer = textoFer & vbCrLf
                textoFer = textoFer & Format$(d, "dd\/mm") & " · " & CStr(fer(CLng(d)))
            End If
        Else
            casa.Ocultar
        End If
    Next i

    lblFeriados.Caption = textoFer
End Sub

Private Sub lblAnterior_Click()
    mMes = mMes - 1
    If mMes < 1 Then
        mMes = 12
        mAno = mAno - 1
    End If
    DesenharCalendario
End Sub

Private Sub lblProximo_Click()
    mMes = mMes + 1
    If mMes > 12 Then
        mMes = 1
        mAno = mAno + 1
    End If
    DesenharCalendario
End Sub

Private Sub lblHoje_Click()
    SelecionarDia Date
End Sub

' Chamado pelas casas do calendário e pela lista de urgentes
Public Sub SelecionarDia(ByVal dia As Date)
    If mEditandoID <> 0 Then LimparCampos
    mDiaSel = dia
    PreencherDias cboDia, mDiaSel
    mAno = Year(dia)
    mMes = Month(dia)
    DesenharCalendario
    CarregarNotas False
End Sub

' ============================================================
'  Lista de anotações do dia
' ============================================================

Private Sub CarregarNotas(ByVal irParaOFim As Boolean)
    Dim notas As Collection, n As clsAnotacao, cartao As clsCartaoNota
    Dim i As Long, topo As Single, rolagem As Single

    Set notas = mRepo.ListarPorData(mDiaSel)

    lblTitNotas.Caption = "ANOTAÇÕES DE " & Format$(mDiaSel, "dd\/mm\/yyyy") & " · " & NomeDiaSemana(mDiaSel)
    If notas.Count = 1 Then
        lblContagem.Caption = "1 anotação · dois cliques abrem"
    Else
        lblContagem.Caption = notas.Count & " anotações · dois cliques abrem"
    End If
    lblVazio.Visible = (notas.Count = 0)

    ' cria cartões que estiverem faltando
    Do While mCartoes.Count < notas.Count
        Set cartao = New clsCartaoNota
        cartao.Criar fraLista, mCartoes.Count + 1
        mCartoes.Add cartao
    Loop

    topo = 6
    For i = 1 To mCartoes.Count
        Set cartao = mCartoes(i)
        If i <= notas.Count Then
            Set n = notas(i)
            topo = topo + cartao.Exibir(n, topo) + 6
        Else
            cartao.Ocultar
        End If
    Next i

    If topo < fraLista.InsideHeight Then topo = fraLista.InsideHeight
    fraLista.ScrollHeight = topo

    If irParaOFim Then
        rolagem = topo - fraLista.InsideHeight
        If rolagem < 0 Then rolagem = 0
        fraLista.ScrollTop = rolagem
    ElseIf fraLista.ScrollTop > topo - fraLista.InsideHeight Then
        fraLista.ScrollTop = 0
    End If
End Sub

Private Sub CarregarUrgentes()
    Dim avisos As Collection, n As clsAnotacao, item As clsItemAviso
    Dim i As Long, topo As Single

    Set avisos = mRepo.ListarAvisos
    lblSemAvisos.Visible = (avisos.Count = 0)

    Do While mItens.Count < avisos.Count
        Set item = New clsItemAviso
        item.Criar fraAvisos, mItens.Count + 1
        mItens.Add item
    Loop

    topo = 6
    For i = 1 To mItens.Count
        Set item = mItens(i)
        If i <= avisos.Count Then
            Set n = avisos(i)
            topo = topo + item.Exibir(n, topo) + 6
        Else
            item.Ocultar
        End If
    Next i

    If topo < fraAvisos.InsideHeight Then topo = fraAvisos.InsideHeight
    fraAvisos.ScrollHeight = topo
    If fraAvisos.ScrollTop > topo - fraAvisos.InsideHeight Then fraAvisos.ScrollTop = 0
End Sub

Private Sub AtualizarTudo(ByVal irParaOFim As Boolean)
    DesenharCalendario
    CarregarNotas irParaOFim
    CarregarUrgentes
End Sub

' ============================================================
'  Ações dos cartões (chamadas por clsCartaoNota)
' ============================================================

Public Sub AoAlternarUrgente(ByVal idNota As Long)
    mRepo.AlternarUrgente idNota
    mRepo.Persistir
    AtualizarTudo False
End Sub

' Quadrado de baixa do cartão: finaliza ou reabre a anotação
Public Sub AoAlternarFinalizada(ByVal idNota As Long)
    mRepo.AlternarFinalizada idNota
    mRepo.Persistir
    AtualizarTudo False
End Sub

Public Sub AoEditar(ByVal idNota As Long)
    AoAbrirNota idNota
End Sub

' Dois cliques no cartão (ou "editar"): abre a anotação inteira em frmNota
Public Sub AoAbrirNota(ByVal idNota As Long)
    If mEditandoID <> 0 Then LimparCampos
    frmNota.Abrir idNota
    Unload frmNota
    AtualizarTudo False
End Sub

Public Sub AoExcluir(ByVal idNota As Long)
    If MsgBox("Excluir esta anotação? Esta ação não pode ser desfeita.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Painel de Anotações") <> vbYes Then Exit Sub
    If mEditandoID = idNota Then LimparCampos
    mRepo.Excluir idNota
    mRepo.Persistir
    AtualizarTudo False
End Sub

' ============================================================
'  Nova anotação / edição
' ============================================================

Private Sub LimparCampos()
    mEditandoID = 0
    txtTitulo.Text = ""
    txtConteudo.Text = ""
    chkUrgente.Value = False
    cboDia.ListIndex = 0
    cboHora.ListIndex = 36
    lblTitNova.Caption = "NOVA ANOTAÇÃO"
    btnSalvar.Caption = "SALVAR ANOTAÇÃO"
    btnCancelar.Visible = False
End Sub

Private Sub btnCancelar_Click()
    LimparCampos
End Sub

Private Sub btnSalvar_Click()
    Dim n As clsAnotacao, titulo As String, conteudo As String, nova As Boolean

    titulo = Trim$(txtTitulo.Text)
    conteudo = txtConteudo.Text

    If Len(titulo) = 0 And Len(Trim$(conteudo)) = 0 Then
        MsgBox "Escreva um título ou um conteúdo antes de salvar.", vbInformation, "Painel de Anotações"
        txtTitulo.SetFocus
        Exit Sub
    End If
    If Len(titulo) = 0 Then titulo = "(sem título)"

    If mEditandoID <> 0 Then Set n = mRepo.Obter(mEditandoID)
    If n Is Nothing Then
        nova = True
        Set n = New clsAnotacao
        n.Data = mDiaSel
        n.Hora = Format$(Now, "hh:nn")
        n.Lembrete = LembreteDosCombos(cboDia, cboHora, mDiaSel)
    End If

    n.Titulo = titulo
    n.Conteudo = conteudo
    n.Urgente = (chkUrgente.Value = True)

    On Error GoTo Falha
    mRepo.Salvar n
    mRepo.Persistir
    On Error GoTo 0

    ' campo volta vazio, pronto para a próxima anotação
    LimparCampos
    AtualizarTudo nova
    txtTitulo.SetFocus
    Exit Sub

Falha:
    MsgBox "Não foi possível gravar a anotação." & vbCrLf & vbCrLf & _
           "Detalhe: " & Err.Description, vbExclamation, "Painel de Anotações"
End Sub
