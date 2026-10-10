VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiaCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Uma casa do calendário. O painel cria 42 destas (6 semanas x 7 dias)
' e apenas troca o conteúdo quando o mês muda.

Private WithEvents mRotulo As MSForms.Label
Attribute mRotulo.VB_VarHelpID = -1
Private WithEvents mMarca As MSForms.Label
Attribute mMarca.VB_VarHelpID = -1
Private WithEvents mAlerta As MSForms.Label
Attribute mAlerta.VB_VarHelpID = -1
Private mDia As Date

Public Sub Criar(ByVal pai As Object, ByVal indice As Long, ByVal esq As Single, ByVal topo As Single)
    Set mRotulo = pai.Controls.Add("Forms.Label.1", "lblDia" & indice, True)
    With mRotulo
        .Left = esq
        .Top = topo
        .Width = 34
        .Height = 22
        .Caption = ""
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
    End With

    Set mMarca = pai.Controls.Add("Forms.Label.1", "lblMarca" & indice, True)
    With mMarca
        .Left = esq + 26
        .Top = topo + 1
        .Width = 5
        .Height = 9
        .Caption = ""
        .BackStyle = fmBackStyleOpaque
        .Visible = False
    End With

    ' exclamação de lembrete ativo, no canto oposto ao da marca
    Set mAlerta = pai.Controls.Add("Forms.Label.1", "lblAlertaDia" & indice, True)
    With mAlerta
        .Left = esq + 1
        .Top = topo + 1
        .Width = 8
        .Height = 11
        .Caption = "!"
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .Font.Bold = True
        .BackStyle = fmBackStyleOpaque
        .Visible = False
    End With
End Sub

' marca: soma de 1 (tem anotação), 2 (tem urgente), 4 (tem lembrete ativo)
' e 8 (tem anotação em aberto; sem o 8, tudo no dia está finalizado)
Public Sub Mostrar(ByVal dia As Date, ByVal selecionado As Boolean, ByVal ehHoje As Boolean, _
                   ByVal ehFeriado As Boolean, ByVal marca As Long)
    Dim fundo As Long, texto As Long, borda As Long, ds As Long

    mDia = dia
    ds = Weekday(dia, vbSunday)

    fundo = RGB(38, 45, 58)
    texto = RGB(232, 236, 241)
    If ds = vbSunday Or ds = vbSaturday Then
        fundo = RGB(28, 33, 42)
        texto = RGB(154, 163, 178)
    End If
    If ehFeriado Then
        fundo = RGB(110, 38, 38)
        texto = RGB(255, 214, 214)
    End If
    borda = fundo
    If ehHoje Then borda = RGB(242, 194, 48)
    If selecionado Then
        fundo = RGB(242, 194, 48)
        texto = RGB(24, 28, 36)
        borda = RGB(242, 194, 48)
    End If

    With mRotulo
        .Caption = CStr(Day(dia))
        .BackColor = fundo
        .ForeColor = texto
        .BorderColor = borda
        .Visible = True
    End With

    With mMarca
        If (marca And 2) <> 0 Then
            .BackColor = RGB(239, 83, 80)
            .Visible = True
        ElseIf (marca And 1) <> 0 And (marca And 8) = 0 Then
            .BackColor = IIf(selecionado, RGB(27, 94, 32), RGB(63, 185, 80))
            .Visible = True
        ElseIf (marca And 1) <> 0 Then
            .BackColor = IIf(selecionado, RGB(24, 28, 36), RGB(200, 206, 216))
            .Visible = True
        Else
            .Visible = False
        End If
    End With

    With mAlerta
        .BackColor = IIf(selecionado, RGB(24, 28, 36), RGB(242, 194, 48))
        .ForeColor = IIf(selecionado, RGB(242, 194, 48), RGB(24, 28, 36))
        .Visible = ((marca And 4) <> 0)
    End With
End Sub

Public Sub Ocultar()
    mRotulo.Visible = False
    mMarca.Visible = False
    mAlerta.Visible = False
End Sub

Private Sub mRotulo_Click()
    frmPainel.SelecionarDia mDia
End Sub

Private Sub mMarca_Click()
    frmPainel.SelecionarDia mDia
End Sub

Private Sub mAlerta_Click()
    frmPainel.SelecionarDia mDia
End Sub
