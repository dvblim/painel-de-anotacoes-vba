VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCartaoNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Um cartão de anotação dentro da lista do dia: quadrado de baixa, hora, título, prévia do
' conteúdo, situação do lembrete, os atalhos "editar" e "excluir" e a fita
' de marca-página (urgente). Dois cliques no cartão abrem a anotação inteira.
' Os cartões são reaproveitados: o painel cria quantos precisar e
' esconde os que sobram, sem nunca remover controles.

Private Const ESQ As Single = 6
Private Const LARG As Single = 588
Private Const ALT_FITA As Single = 34
Private Const ALT_MAX_CORPO As Single = 40     ' prévia de cerca de 3 linhas

Private WithEvents mFita As MSForms.Label
Attribute mFita.VB_VarHelpID = -1
Private WithEvents mEntalhe As MSForms.Label
Attribute mEntalhe.VB_VarHelpID = -1
Private WithEvents mEditar As MSForms.Label
Attribute mEditar.VB_VarHelpID = -1
Private WithEvents mExcluir As MSForms.Label
Attribute mExcluir.VB_VarHelpID = -1
Private WithEvents mCheck As MSForms.Label
Attribute mCheck.VB_VarHelpID = -1
Private WithEvents mFundo As MSForms.Label
Attribute mFundo.VB_VarHelpID = -1
Private WithEvents mAlerta As MSForms.Label
Attribute mAlerta.VB_VarHelpID = -1
Private WithEvents mHora As MSForms.Label
Attribute mHora.VB_VarHelpID = -1
Private WithEvents mTitulo As MSForms.Label
Attribute mTitulo.VB_VarHelpID = -1
Private WithEvents mTituloRiscado As MSForms.Label
Attribute mTituloRiscado.VB_VarHelpID = -1
Private WithEvents mInfo As MSForms.Label
Attribute mInfo.VB_VarHelpID = -1
Private WithEvents mCorpo As MSForms.Label
Attribute mCorpo.VB_VarHelpID = -1

Private mID As Long

Private Function Rotulo(ByVal pai As Object, ByVal nome As String, ByVal tam As Single, _
                        ByVal negrito As Boolean, ByVal cor As Long) As MSForms.Label
    Dim c As MSForms.Label
    Set c = pai.Controls.Add("Forms.Label.1", nome, True)
    With c
        .Caption = ""
        .Font.Name = "Segoe UI"
        .Font.Size = tam
        .Font.Bold = negrito
        .ForeColor = cor
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
    End With
    Set Rotulo = c
End Function

Public Sub Criar(ByVal pai As Object, ByVal indice As Long)
    Dim sufixo As String
    sufixo = CStr(indice)

    ' o fundo é criado primeiro para ficar atrás dos demais
    Set mFundo = Rotulo(pai, "crtFundo" & sufixo, 9, False, RGB(232, 236, 241))
    With mFundo
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(42, 49, 63)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(58, 67, 84)
        .ControlTipText = "Dois cliques para abrir a anotação"
    End With

    ' quadrado de baixa: vazio = em aberto, verde com check = finalizada
    Set mCheck = Rotulo(pai, "crtCheck" & sufixo, 9, True, RGB(16, 21, 28))
    With mCheck
        .Font.Name = "Segoe UI Symbol"
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
    End With

    ' exclamação do lembrete ativo
    Set mAlerta = Rotulo(pai, "crtAlerta" & sufixo, 13, True, RGB(24, 28, 36))
    With mAlerta
        .Caption = "!"
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(242, 194, 48)
    End With

    Set mHora = Rotulo(pai, "crtHora" & sufixo, 8, True, RGB(154, 163, 178))
    Set mTitulo = Rotulo(pai, "crtTitulo" & sufixo, 10, True, RGB(232, 236, 241))
    ' título riscado, usado só depois da baixa (um rótulo à parte: a propriedade
    ' de riscado nunca é alterada no título normal)
    Set mTituloRiscado = Rotulo(pai, "crtTituloRiscado" & sufixo, 10, True, RGB(180, 189, 202))
    mTituloRiscado.Font.Strikethrough = True
    Set mInfo = Rotulo(pai, "crtInfo" & sufixo, 8, True, RGB(154, 163, 178))
    Set mCorpo = Rotulo(pai, "crtCorpo" & sufixo, 9, False, RGB(232, 236, 241))
    mCorpo.WordWrap = True

    Set mEditar = Rotulo(pai, "crtEditar" & sufixo, 8, False, RGB(125, 170, 240))
    mEditar.Caption = "editar"
    mEditar.Font.Underline = True

    Set mExcluir = Rotulo(pai, "crtExcluir" & sufixo, 8, False, RGB(240, 128, 128))
    mExcluir.Caption = "excluir"
    mExcluir.Font.Underline = True

    ' fita do marca-página: um retângulo com um entalhe em "V" na ponta
    Set mFita = Rotulo(pai, "crtFita" & sufixo, 8, False, RGB(255, 255, 255))
    mFita.BackStyle = fmBackStyleOpaque

    Set mEntalhe = Rotulo(pai, "crtEntalhe" & sufixo, 16, False, RGB(42, 49, 63))
    With mEntalhe
        .Font.Name = "Segoe UI Symbol"
        .Caption = ChrW(&H25B2)
        .TextAlign = fmTextAlignCenter
    End With

    Ocultar
End Sub

' Posiciona e preenche o cartão; devolve a altura ocupada
Public Function Exibir(ByVal n As clsAnotacao, ByVal topo As Single) As Single
    Dim altCorpo As Single, alt As Single, texto As String
    Dim estado As Long, ativo As Boolean, temInfo As Boolean
    Dim desloc As Single, largCorpo As Single, yCorpo As Single
    Dim corFundo As Long, corBorda As Long, corInfo As Long, corTexto As Long
    Const RECUO As Single = 22       ' espaço do quadrado de baixa

    mID = n.ID
    texto = n.Conteudo
    estado = n.EstadoLembrete
    ativo = (estado = 2 Or estado = 3)
    temInfo = (estado >= 1) Or n.Finalizada

    ' cores conforme a situação do lembrete
    corFundo = RGB(42, 49, 63)
    corBorda = RGB(58, 67, 84)
    corInfo = RGB(154, 163, 178)
    corTexto = RGB(232, 236, 241)
    If n.Finalizada Then
        corFundo = RGB(31, 51, 39)
        corBorda = RGB(46, 125, 79)
        corInfo = RGB(111, 208, 140)
        corTexto = RGB(180, 189, 202)
    ElseIf estado = 2 Then
        corFundo = RGB(58, 51, 32)
        corBorda = RGB(242, 194, 48)
        corInfo = RGB(242, 194, 48)
    ElseIf estado = 3 Then
        corFundo = RGB(74, 36, 36)
        corBorda = RGB(239, 83, 80)
        corInfo = RGB(255, 170, 170)
    End If

    If ativo Then desloc = 24 Else desloc = 0
    desloc = desloc + RECUO
    largCorpo = LARG - 64 - desloc

    ' mede a altura do conteúdo e limita à prévia
    If Len(Trim$(texto)) > 0 Then
        With mCorpo
            .AutoSize = False
            .Width = largCorpo
            .Height = 12
            .Caption = texto
            .AutoSize = True
            altCorpo = .Height
            .AutoSize = False
            If altCorpo > ALT_MAX_CORPO Then altCorpo = ALT_MAX_CORPO
            .Width = largCorpo
            .Height = altCorpo
        End With
    Else
        mCorpo.Caption = ""
        altCorpo = 0
    End If

    If temInfo Then yCorpo = 42 Else yCorpo = 28
    alt = yCorpo + altCorpo + 10
    If alt < 48 Then alt = 48

    With mFundo
        .Left = ESQ: .Top = topo: .Width = LARG: .Height = alt
        .BackColor = corFundo
        .BorderColor = corBorda
        .Visible = True
    End With
    With mCheck
        .Left = ESQ + 8: .Top = topo + 8: .Width = 14: .Height = 14
        If n.Finalizada Then
            .Caption = ChrW(&H2713)
            .BackColor = RGB(63, 185, 80)
            .BorderColor = RGB(63, 185, 80)
            .ControlTipText = "Reabrir (desfazer a baixa)"
        Else
            .Caption = ""
            .BackColor = corFundo
            .BorderColor = RGB(141, 151, 168)
            .ControlTipText = "Dar baixa: marcar como finalizada"
        End If
        .Visible = True
    End With
    With mAlerta
        .Left = ESQ + 8 + RECUO: .Top = topo + 8: .Width = 16: .Height = 22
        .BackColor = corBorda
        .Visible = ativo
    End With
    With mHora
        .Left = ESQ + 8 + desloc: .Top = topo + 8: .Width = 34: .Height = 12
        .Caption = n.Hora
        .Visible = True
    End With
    With mTitulo
        .Left = ESQ + 44 + desloc: .Top = topo + 6: .Width = 410 - desloc: .Height = 16
        .Caption = n.Titulo
        .Visible = Not n.Finalizada
    End With
    With mTituloRiscado
        .Left = ESQ + 44 + desloc: .Top = topo + 6: .Width = 410 - desloc: .Height = 16
        .Caption = n.Titulo
        .Visible = n.Finalizada
    End With
    With mInfo
        .Left = ESQ + 8 + desloc: .Top = topo + 25: .Width = largCorpo: .Height = 12
        .Caption = n.TextoSituacao
        .ForeColor = corInfo
        .Visible = temInfo
    End With
    With mEditar
        .Left = ESQ + LARG - 124: .Top = topo + 8: .Width = 32: .Height = 12
        .Visible = True
    End With
    With mExcluir
        .Left = ESQ + LARG - 88: .Top = topo + 8: .Width = 36: .Height = 12
        .Visible = True
    End With
    With mCorpo
        .Left = ESQ + 8 + desloc: .Top = topo + yCorpo
        .ForeColor = corTexto
        .Visible = (altCorpo > 0)
    End With
    With mFita
        .Left = ESQ + LARG - 36: .Top = topo: .Width = 18: .Height = ALT_FITA
        .BackColor = IIf(n.Urgente, RGB(239, 83, 80), RGB(85, 94, 112))
        .ControlTipText = IIf(n.Urgente, "Remover marcador de urgente", "Marcar como urgente")
        .Visible = True
    End With
    With mEntalhe
        .Left = ESQ + LARG - 36: .Top = topo + ALT_FITA - 16: .Width = 18: .Height = 24
        .ForeColor = corFundo
        .ControlTipText = mFita.ControlTipText
        .Visible = True
    End With

    Exibir = alt
End Function

Public Sub Ocultar()
    mFundo.Visible = False
    mCheck.Visible = False
    mAlerta.Visible = False
    mHora.Visible = False
    mTitulo.Visible = False
    mTituloRiscado.Visible = False
    mInfo.Visible = False
    mCorpo.Visible = False
    mEditar.Visible = False
    mExcluir.Visible = False
    mFita.Visible = False
    mEntalhe.Visible = False
    mID = 0
End Sub

Private Sub Abrir()
    If mID <> 0 Then frmPainel.AoAbrirNota mID
End Sub

Private Sub mCheck_Click()
    If mID <> 0 Then frmPainel.AoAlternarFinalizada mID
End Sub

Private Sub mFita_Click()
    If mID <> 0 Then frmPainel.AoAlternarUrgente mID
End Sub

Private Sub mEntalhe_Click()
    If mID <> 0 Then frmPainel.AoAlternarUrgente mID
End Sub

Private Sub mEditar_Click()
    Abrir
End Sub

Private Sub mExcluir_Click()
    If mID <> 0 Then frmPainel.AoExcluir mID
End Sub

Private Sub mFundo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mAlerta_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mHora_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mTitulo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mTituloRiscado_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mInfo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub

Private Sub mCorpo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Abrir
End Sub
