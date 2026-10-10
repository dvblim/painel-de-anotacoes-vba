VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNota
   Caption         =   "Anotação"
   ClientHeight    =   8900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Janela aberta com dois cliques em uma anotação do painel:
' mostra o conteúdo inteiro e permite editar, definir o lembrete,
' dar baixa (ou reabrir) ou excluir.

Private mRepo As clsRepositorio
Private mNota As clsAnotacao

Private mCorTexto As Long
Private mCorCinza As Long

Private WithEvents btnSalvar As MSForms.CommandButton
Attribute btnSalvar.VB_VarHelpID = -1
Private WithEvents btnBaixa As MSForms.CommandButton
Attribute btnBaixa.VB_VarHelpID = -1
Private WithEvents btnExcluir As MSForms.CommandButton
Attribute btnExcluir.VB_VarHelpID = -1
Private WithEvents btnFechar As MSForms.CommandButton
Attribute btnFechar.VB_VarHelpID = -1

Private lblCab As MSForms.Label
Private lblStatus As MSForms.Label
Private txtTitulo As MSForms.TextBox
Private txtConteudo As MSForms.TextBox
Private cboDia As MSForms.ComboBox
Private cboHora As MSForms.ComboBox
Private chkUrgente As MSForms.CheckBox

Private Function NovoRotulo(ByVal nome As String, ByVal texto As String, ByVal esq As Single, ByVal topo As Single, _
                            ByVal larg As Single, ByVal alt As Single, Optional ByVal tam As Single = 9, _
                            Optional ByVal negrito As Boolean = False) As MSForms.Label
    Dim c As MSForms.Label
    Set c = Me.Controls.Add("Forms.Label.1", nome, True)
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

Private Function NovoBotao(ByVal nome As String, ByVal texto As String, ByVal esq As Single, ByVal larg As Single, _
                           ByVal fundo As Long, ByVal cor As Long, ByVal negrito As Boolean) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", nome, True)
    With b
        .Caption = texto
        .Left = esq
        .Top = 400
        .Width = larg
        .Height = 26
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = negrito
        .BackColor = fundo
        .ForeColor = cor
    End With
    Set NovoBotao = b
End Function

Private Function NovoCombo(ByVal nome As String, ByVal esq As Single, ByVal larg As Single) As MSForms.ComboBox
    Dim c As MSForms.ComboBox
    Set c = Me.Controls.Add("Forms.ComboBox.1", nome, True)
    With c
        .Left = esq
        .Top = 60
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

Private Sub UserForm_Initialize()
    Dim c As MSForms.Label

    mCorTexto = RGB(232, 236, 241)
    mCorCinza = RGB(154, 163, 178)

    Me.Caption = "Anotação"
    Me.BackColor = RGB(24, 28, 36)
    Me.Width = 640
    Me.Height = 466

    Set lblCab = NovoRotulo("lblCab", "", 12, 10, 604, 12, 8, True)
    lblCab.ForeColor = mCorCinza
    Set lblStatus = NovoRotulo("lblStatus", "", 12, 24, 604, 14, 9, True)

    Set c = NovoRotulo("lblCampoTitulo", "Título", 12, 44, 100, 12, 8, True)
    c.ForeColor = mCorCinza
    Set txtTitulo = Me.Controls.Add("Forms.TextBox.1", "txtTitulo", True)
    With txtTitulo
        .Left = 12
        .Top = 58
        .Width = 320
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(74, 84, 104)
        .BackColor = RGB(32, 38, 49)
        .ForeColor = mCorTexto
        .MaxLength = 200
    End With

    Set c = NovoRotulo("lblCampoConteudo", "Conteúdo", 12, 86, 100, 12, 8, True)
    c.ForeColor = mCorCinza
    Set txtConteudo = Me.Controls.Add("Forms.TextBox.1", "txtConteudo", True)
    With txtConteudo
        .Left = 12
        .Top = 100
        .Width = 604
        .Height = 286
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .MultiLine = True
        .WordWrap = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(74, 84, 104)
        .BackColor = RGB(32, 38, 49)
        .ForeColor = mCorTexto
    End With

    ' lembrete na linha do título: as listas abrem sobre a caixa de conteúdo
    Set c = NovoRotulo("lblLembrarEm", "Lembrar em", 342, 44, 100, 12, 8, True)
    c.ForeColor = RGB(242, 194, 48)
    Set cboDia = NovoCombo("cboDia", 342, 190)
    cboDia.ColumnCount = 2
    cboDia.ColumnWidths = "180 pt;0 pt"
    Set c = NovoRotulo("lblHorario", "Horário", 542, 44, 74, 12, 8, True)
    c.ForeColor = RGB(242, 194, 48)
    Set cboHora = NovoCombo("cboHora", 542, 74)

    Set chkUrgente = Me.Controls.Add("Forms.CheckBox.1", "chkUrgente", True)
    With chkUrgente
        .Left = 12
        .Top = 404
        .Width = 150
        .Height = 18
        .Caption = "Marcar como urgente"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = RGB(24, 28, 36)
        .ForeColor = mCorTexto
        .Value = False
    End With

    Set btnBaixa = NovoBotao("btnBaixa", "Dar baixa", 216, 100, RGB(63, 185, 80), RGB(16, 21, 28), True)
    Set btnExcluir = NovoBotao("btnExcluir", "Excluir", 322, 70, RGB(58, 67, 84), RGB(240, 128, 128), False)
    Set btnFechar = NovoBotao("btnFechar", "Fechar", 398, 70, RGB(58, 67, 84), mCorTexto, False)
    Set btnSalvar = NovoBotao("btnSalvar", "SALVAR ALTERAÇÕES", 474, 142, RGB(242, 194, 48), RGB(24, 28, 36), True)
End Sub

' Carrega a anotação e mostra a janela (chamado pelo painel)
Public Sub Abrir(ByVal idNota As Long)
    Dim estado As Long

    Set mRepo = New clsRepositorio
    Set mNota = mRepo.Obter(idNota)
    If mNota Is Nothing Then Exit Sub

    lblCab.Caption = "ANOTAÇÃO DE " & Format$(mNota.Data, "dd\/mm\/yyyy") & " · " & NomeDia(mNota.Data) & _
                     " · registrada às " & mNota.Hora
    estado = mNota.EstadoLembrete
    lblStatus.Caption = mNota.TextoSituacao
    Select Case estado
        Case 2: lblStatus.ForeColor = RGB(242, 194, 48)
        Case 3: lblStatus.ForeColor = RGB(255, 170, 170)
        Case Else: lblStatus.ForeColor = mCorCinza
    End Select
    If mNota.Finalizada Then lblStatus.ForeColor = RGB(111, 208, 140)

    txtTitulo.Text = mNota.Titulo
    txtConteudo.Text = mNota.Conteudo
    chkUrgente.Value = mNota.Urgente
    SelecionarCombos cboDia, cboHora, mNota.Data, mNota.Lembrete
    btnBaixa.Caption = IIf(mNota.Finalizada, "Reabrir", "Dar baixa")

    Me.Show
End Sub

' Valida os campos e passa o que foi digitado para a anotação.
' Devolve False se o usuário desistiu ou falta preencher algo.
Private Function AplicarCampos() As Boolean
    Dim titulo As String, conteudo As String, novoLembrete As Date

    titulo = Trim$(txtTitulo.Text)
    conteudo = txtConteudo.Text
    If Len(titulo) = 0 And Len(Trim$(conteudo)) = 0 Then
        MsgBox "Escreva um título ou um conteúdo antes de salvar.", vbInformation, "Painel de Anotações"
        Exit Function
    End If
    If Len(titulo) = 0 Then titulo = "(sem título)"

    novoLembrete = LembreteDosCombos(cboDia, cboHora, mNota.Data)
    If mNota.Lembrete <> 0 And novoLembrete = 0 Then
        If MsgBox("Deseja mesmo excluir o lembrete desta anotação?", _
                  vbYesNo + vbQuestion + vbDefaultButton2, "Painel de Anotações") <> vbYes Then Exit Function
    End If
    ' lembrete alterado volta a valer, mesmo que o anterior tivesse sido concluído
    If novoLembrete <> mNota.Lembrete Then mNota.LembreteConcluido = False

    mNota.Titulo = titulo
    mNota.Conteudo = conteudo
    mNota.Urgente = (chkUrgente.Value = True)
    mNota.Lembrete = novoLembrete
    AplicarCampos = True
End Function

Private Sub Gravar()
    On Error GoTo Falha
    mRepo.Salvar mNota
    mRepo.Persistir
    Me.Hide
    Exit Sub
Falha:
    MsgBox "Não foi possível gravar a anotação." & vbCrLf & vbCrLf & _
           "Detalhe: " & Err.Description, vbExclamation, "Painel de Anotações"
End Sub

Private Sub btnSalvar_Click()
    If Not AplicarCampos Then Exit Sub
    Gravar
End Sub

' Dá baixa (ou reabre) e grava também o que foi editado na janela
Private Sub btnBaixa_Click()
    If Not AplicarCampos Then Exit Sub
    mNota.Finalizada = Not mNota.Finalizada
    If mNota.Finalizada Then mNota.FinalizadaEm = Now Else mNota.FinalizadaEm = 0
    Gravar
End Sub

Private Sub btnExcluir_Click()
    If MsgBox("Excluir esta anotação? Esta ação não pode ser desfeita.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Painel de Anotações") <> vbYes Then Exit Sub
    mRepo.Excluir mNota.ID
    mRepo.Persistir
    Me.Hide
End Sub

Private Sub btnFechar_Click()
    Me.Hide
End Sub
