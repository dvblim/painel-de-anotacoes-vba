VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Uma anotação (ou ata) registrada em uma data.
' ID = 0 indica anotação nova, ainda não gravada na base.

Public ID As Long
Public Data As Date
Public Hora As String
Public Titulo As String
Public Conteudo As String
Public Urgente As Boolean
Public Lembrete As Date               ' dia e hora do lembrete; 0 = sem lembrete
Public LembreteConcluido As Boolean
Public Finalizada As Boolean          ' baixa dada na atividade ou compromisso
Public FinalizadaEm As Date

' 0 = sem lembrete, 1 = programado (ainda não chegou a hora),
' 2 = ativo, 3 = ativo e a data da anotação já passou, 4 = concluído
Public Function EstadoLembrete() As Long
    If Lembrete = 0 Or Finalizada Then
        EstadoLembrete = 0
    ElseIf LembreteConcluido Then
        EstadoLembrete = 4
    ElseIf Now < Lembrete Then
        EstadoLembrete = 1
    ElseIf Date > Data Then
        EstadoLembrete = 3
    Else
        EstadoLembrete = 2
    End If
End Function

Public Function LembreteAtivo() As Boolean
    Dim e As Long
    e = EstadoLembrete
    LembreteAtivo = (e = 2 Or e = 3)
End Function

Public Function TextoLembrete() As String
    Dim quando As String
    If Lembrete = 0 Then Exit Function
    quando = Format$(Lembrete, "dd\/mm") & " às " & Format$(Lembrete, "hh:nn")
    Select Case EstadoLembrete
        Case 1: TextoLembrete = "Lembrete programado para " & quando
        Case 2: TextoLembrete = "Lembrete ativo desde " & quando
        Case 3: TextoLembrete = "Lembrete ativo desde " & quando & " · a data já passou"
        Case 4: TextoLembrete = "Lembrete concluído"
    End Select
End Function

' Linha de situação mostrada no cartão e na janela da anotação
Public Function TextoSituacao() As String
    If Finalizada Then
        If FinalizadaEm > 0 Then
            TextoSituacao = "Finalizada em " & Format$(FinalizadaEm, "dd\/mm") & " às " & Format$(FinalizadaEm, "hh:nn")
        Else
            TextoSituacao = "Finalizada"
        End If
    Else
        TextoSituacao = TextoLembrete
    End If
End Function
