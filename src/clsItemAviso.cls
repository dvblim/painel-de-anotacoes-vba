VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsItemAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Uma linha da lista "Alertas e urgentes": marcador + data e título completos,
' com quebra de linha quando o título é longo. Clicar leva à data da anotação.
' Assim como os cartões, os itens são reaproveitados e nunca removidos.

Private Const LARG_TEXTO As Single = 212

Private WithEvents mMarca As MSForms.Label
Attribute mMarca.VB_VarHelpID = -1
Private WithEvents mTexto As MSForms.Label
Attribute mTexto.VB_VarHelpID = -1
Private mData As Date

Public Sub Criar(ByVal pai As Object, ByVal indice As Long)
    Set mMarca = pai.Controls.Add("Forms.Label.1", "avsMarca" & indice, True)
    With mMarca
        .Caption = ""
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = RGB(24, 28, 36)
        .BackStyle = fmBackStyleOpaque
        .Visible = False
    End With

    Set mTexto = pai.Controls.Add("Forms.Label.1", "avsTexto" & indice, True)
    With mTexto
        .Caption = ""
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(232, 236, 241)
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .ControlTipText = "Clique para ir à data da anotação"
        .Visible = False
    End With
End Sub

' Posiciona e preenche o item; devolve a altura ocupada
Public Function Exibir(ByVal n As clsAnotacao, ByVal topo As Single) As Single
    Dim alt As Single

    mData = n.Data

    With mTexto
        .AutoSize = False
        .Width = LARG_TEXTO
        .Height = 12
        .Caption = Format$(n.Data, "dd\/mm\/yyyy") & " · " & n.Titulo
        .AutoSize = True
        alt = .Height
        .AutoSize = False
        .Width = LARG_TEXTO
        .Height = alt
        .Left = 20
        .Top = topo
        .Visible = True
    End With
    If alt < 13 Then alt = 13

    With mMarca
        If n.LembreteAtivo Then
            ' exclamação: âmbar, ou vermelha quando a data já passou
            .Caption = "!"
            .Left = 5: .Top = topo + 1: .Width = 10: .Height = 12
            .BackColor = IIf(n.EstadoLembrete = 3, RGB(239, 83, 80), RGB(242, 194, 48))
        Else
            ' urgente: pequena fita vermelha
            .Caption = ""
            .Left = 7: .Top = topo + 2: .Width = 6: .Height = 10
            .BackColor = RGB(239, 83, 80)
        End If
        .Visible = True
    End With

    Exibir = alt
End Function

Public Sub Ocultar()
    mMarca.Visible = False
    mTexto.Visible = False
End Sub

Private Sub mMarca_Click()
    frmPainel.SelecionarDia mData
End Sub

Private Sub mTexto_Click()
    frmPainel.SelecionarDia mData
End Sub
