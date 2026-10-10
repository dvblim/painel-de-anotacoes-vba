"""Gera o InstalarPainel.bas a partir dos arquivos de src/.

Uso (na raiz do repositório):  python ferramentas/gerar_instalador.py

Os arquivos de src/ e o instalador ficam em Windows-1252 com CRLF,
que é o formato que o editor do VBA importa sem estragar os acentos.
"""
import pathlib, re

RAIZ = pathlib.Path(__file__).resolve().parent.parent
SRC = RAIZ / 'src'
OUT = RAIZ


def body(name):
    """Código do componente sem o cabeçalho de exportação e sem as linhas Attribute."""
    t = (SRC / name).read_bytes().decode('cp1252').replace('\r\n', '\n')
    if 'Attribute VB_Exposed = False\n' in t:
        t = t.split('Attribute VB_Exposed = False\n', 1)[1]
    lines = [l for l in t.split('\n') if not l.startswith('Attribute ')]
    while lines and lines[-1] == '':
        lines.pop()
    return lines


comps = [('frmPainel','frmPainel.frm',3),('frmNota','frmNota.frm',3),('modPainel','modPainel.bas',1),('clsAnotacao','clsAnotacao.cls',2),
         ('clsRepositorio','clsRepositorio.cls',2),('clsDiaCalendario','clsDiaCalendario.cls',2),('clsCartaoNota','clsCartaoNota.cls',2),('clsItemAviso','clsItemAviso.cls',2)]
gen = []; check = {}
for comp, fn, tipo in comps:
    ls = body(fn); check[comp] = ls
    chunks = [ls[i:i+90] for i in range(0, len(ls), 90)]
    gen.append(f"Private Function Codigo_{comp}() As String\n    mBuf = \"\"\n" + "".join(f"    {comp}_{i+1}\n" for i in range(len(chunks))) + f"    Codigo_{comp} = mBuf\n    mBuf = \"\"\nEnd Function\n")
    for i, ch in enumerate(chunks):
        gen.append(f"Private Sub {comp}_{i+1}()\n" + "".join('    L "' + l.replace('"','""') + '"\n' for l in ch) + "End Sub\n")
head = r'''Option Explicit

' ============================================================
'  INSTALADOR DO PAINEL DE ANOTAÇÕES
' ------------------------------------------------------------
'  1. Abra uma pasta de trabalho nova (ela vai virar o painel).
'  2. Alt+F11 > Inserir > Módulo, e cole todo este texto
'     (ou arraste este arquivo .bas para dentro do projeto).
'  3. Clique dentro de InstalarPainel e pressione F5.
'
'  Requisito: em Arquivo > Opções > Central de Confiabilidade >
'  Configurações da Central de Confiabilidade > Configurações de
'  Macro, a opção "Confiar no acesso ao modelo de objeto do
'  projeto do VBA" precisa estar marcada.
'
'  Pode ser executado de novo: recria o código, mas não apaga
'  as anotações já gravadas na aba Base.
' ============================================================

Private mBuf As String

Public Sub InstalarPainel()
    Dim proj As Object
    Dim etapa As String

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    If Err.Number <> 0 Or proj Is Nothing Then
        On Error GoTo 0
        MsgBox "O Excel bloqueou o acesso ao projeto VBA." & vbCrLf & vbCrLf & _
               "Abra Arquivo > Opções > Central de Confiabilidade > Configurações da Central de " & _
               "Confiabilidade > Configurações de Macro, marque ""Confiar no acesso ao modelo de " & _
               "objeto do projeto do VBA"" e execute InstalarPainel de novo.", vbExclamation, "Instalador do Painel"
        Exit Sub
    End If
    On Error GoTo Falha

    If proj.Protection <> 0 Then
        MsgBox "O projeto VBA desta pasta de trabalho está protegido por senha. Desbloqueie e tente de novo.", _
               vbExclamation, "Instalador do Painel"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    etapa = "criar as abas e tabelas"
    CriarAbaPainel
    CriarAbaBase
    CriarAbaFeriados

    etapa = "criar o formulário frmPainel"
    InstalarComponente proj, "frmPainel", 3, Codigo_frmPainel
    On Error Resume Next
    proj.References.AddFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0
    On Error GoTo Falha

    etapa = "criar o formulário frmNota"
    InstalarComponente proj, "frmNota", 3, Codigo_frmNota

    etapa = "criar o módulo modPainel"
    InstalarComponente proj, "modPainel", 1, Codigo_modPainel
    etapa = "criar a classe clsAnotacao"
    InstalarComponente proj, "clsAnotacao", 2, Codigo_clsAnotacao
    etapa = "criar a classe clsRepositorio"
    InstalarComponente proj, "clsRepositorio", 2, Codigo_clsRepositorio
    etapa = "criar a classe clsDiaCalendario"
    InstalarComponente proj, "clsDiaCalendario", 2, Codigo_clsDiaCalendario
    etapa = "criar a classe clsCartaoNota"
    InstalarComponente proj, "clsCartaoNota", 2, Codigo_clsCartaoNota

    etapa = "criar a classe clsItemAviso"
    InstalarComponente proj, "clsItemAviso", 2, Codigo_clsItemAviso

    etapa = "criar o botão ABRIR PAINEL"
    CriarBotao

    ThisWorkbook.Worksheets("Painel").Activate
    Application.ScreenUpdating = True

    etapa = "salvar como .xlsm"
    SalvarComoXlsm

    MsgBox "Instalação concluída." & vbCrLf & vbCrLf & _
           "Use o botão ABRIR PAINEL na aba Painel." & vbCrLf & _
           "Este módulo instalador já pode ser removido do projeto.", vbInformation, "Instalador do Painel"
    Exit Sub

Falha:
    Application.ScreenUpdating = True
    MsgBox "A instalação parou na etapa: " & etapa & vbCrLf & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation, "Instalador do Painel"
End Sub

Private Sub L(ByVal texto As String)
    mBuf = mBuf & texto & vbCrLf
End Sub

' Cria o componente (1 = módulo, 2 = classe, 3 = UserForm) ou reaproveita o existente
Private Sub InstalarComponente(ByVal proj As Object, ByVal nome As String, ByVal tipo As Long, ByVal codigo As String)
    Dim comp As Object

    On Error Resume Next
    Set comp = proj.VBComponents(nome)
    On Error GoTo 0

    If comp Is Nothing Then
        Set comp = proj.VBComponents.Add(tipo)
        comp.Name = nome
    End If

    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString codigo
    End With
End Sub

Private Function Aba(ByVal nome As String, ByRef criada As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nome
        criada = True
    End If
    Set Aba = ws
End Function

Private Function TabelaExiste(ByVal ws As Worksheet, ByVal nome As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(nome)
    On Error GoTo 0
    TabelaExiste = Not lo Is Nothing
End Function

Private Sub FormatarCabecalho(ByVal rng As Range)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(43, 52, 69)
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub

Private Sub CriarAbaPainel()
    Dim ws As Worksheet, criada As Boolean
    Set ws = Aba("Painel", criada)
    If Not criada Then Exit Sub

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 110
    ws.Columns("B").WrapText = True
    ws.Columns("B").VerticalAlignment = xlTop

    With ws.Range("B2")
        .Value = "PAINEL DE ANOTAÇÕES"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(43, 52, 69)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .RowHeight = 32
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = RGB(242, 194, 48)
    End With
    ws.Rows("4:5").RowHeight = 24

    ws.Range("B8").Value = "COMO USAR"
    ws.Range("B8").Font.Bold = True
    ws.Range("B8").Font.Size = 11
    ws.Range("B9").Value = "Clique em ABRIR PAINEL (ou Alt+F8 > AbrirPainel). Escolha a data no calendário, " & _
        "escreva título e conteúdo e salve: o campo volta vazio para a próxima anotação."
    ws.Range("B10").Value = "A fita no canto de cada anotação é o marca-página: clique para marcar ou desmarcar " & _
        "o tema como urgente. O quadrado à esquerda dá baixa na anotação. Dois cliques abrem o conteúdo inteiro."
    ws.Range("B11").Value = "As abas Base e Feriados são apenas a base de dados. Alimente as anotações sempre " & _
        "pelo painel; os feriados podem ser editados direto na aba Feriados."

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Sub CriarAbaBase()
    Dim ws As Worksheet, criada As Boolean, lo As ListObject
    Set ws = Aba("Base", criada)
    If TabelaExiste(ws, "tbAnotacoes") Then Exit Sub

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Range("A1:L1").Value = Array("ID", "Data", "Hora", "Título", "Conteúdo", "Urgente", "Registrado em", _
        "Alterado em", "Lembrete", "Lembrete concluído", "Finalizada", "Finalizada em")

    ' linha de exemplo: mostra o formato e pode ser excluída pelo painel
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = Date
    ws.Range("C2").Value = TimeSerial(Hour(Now), Minute(Now), 0)
    ws.Range("D2").Value = "Exemplo · pode excluir pelo painel"
    ws.Range("E2").Value = "Esta linha mostra o formato de cada registro." & vbLf & _
        "Use ""excluir"" no cartão desta anotação, dentro do painel, para removê-la."
    ws.Range("F2").Value = "NÃO"
    ws.Range("G2").Value = Now

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:L2"), , xlYes)
    lo.Name = "tbAnotacoes"
    lo.TableStyle = "TableStyleLight1"
    FormatarCabecalho ws.Range("A1:L1")

    ws.Columns("A").ColumnWidth = 7
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 36
    ws.Columns("E").ColumnWidth = 70
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G:I").ColumnWidth = 18
    ws.Columns("J").ColumnWidth = 20
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 18
    ws.Columns("L").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("B").NumberFormat = "dd/mm/yyyy"
    ws.Columns("C").NumberFormat = "hh:mm"
    ws.Columns("G:I").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("E").WrapText = True
    ws.Columns("A:L").VerticalAlignment = xlTop
End Sub

Private Sub Feriado(ByVal ws As Worksheet, ByRef lin As Long, ByVal ano As Long, ByVal mes As Long, _
                    ByVal dia As Long, ByVal descricao As String)
    lin = lin + 1
    ws.Cells(lin, 1).Value = DateSerial(ano, mes, dia)
    ws.Cells(lin, 2).Value = descricao
End Sub

Private Sub CriarAbaFeriados()
    Dim ws As Worksheet, criada As Boolean, lo As ListObject, lin As Long, ano As Long
    Set ws = Aba("Feriados", criada)
    If TabelaExiste(ws, "tbFeriados") Then Exit Sub

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Range("A1:B1").Value = Array("Data", "Descrição")
    lin = 1

    ' feriados nacionais; a Sexta-feira Santa vem da Páscoa (05/04/2026 e 28/03/2027)
    For ano = 2026 To 2027
        Feriado ws, lin, ano, 1, 1, "Confraternização Universal"
        If ano = 2026 Then Feriado ws, lin, 2026, 4, 3, "Sexta-feira Santa"
        If ano = 2027 Then Feriado ws, lin, 2027, 3, 26, "Sexta-feira Santa"
        Feriado ws, lin, ano, 4, 21, "Tiradentes"
        Feriado ws, lin, ano, 5, 1, "Dia do Trabalho"
        Feriado ws, lin, ano, 9, 7, "Independência do Brasil"
        Feriado ws, lin, ano, 10, 12, "Nossa Senhora Aparecida"
        Feriado ws, lin, ano, 11, 2, "Finados"
        Feriado ws, lin, ano, 11, 15, "Proclamação da República"
        Feriado ws, lin, ano, 11, 20, "Dia da Consciência Negra"
        Feriado ws, lin, ano, 12, 25, "Natal"
    Next ano

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B" & lin), , xlYes)
    lo.Name = "tbFeriados"
    lo.TableStyle = "TableStyleLight1"
    FormatarCabecalho ws.Range("A1:B1")

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 60
    With ws.Range("D1")
        .Value = "Feriados nacionais de 2026 e 2027. Acrescente os estaduais, municipais e os dos " & _
                 "próximos anos em novas linhas da tabela."
        .Font.Italic = True
        .Font.Size = 9
        .WrapText = True
    End With
End Sub

Private Sub CriarBotao()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Worksheets("Painel")

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

' O código só fica guardado em arquivo .xlsm (formato 52)
Private Sub SalvarComoXlsm()
    Dim caminho As Variant

    If ThisWorkbook.FileFormat = 52 And Len(ThisWorkbook.Path) > 0 Then
        ThisWorkbook.Save
        Exit Sub
    End If

    caminho = Application.GetSaveAsFilename(InitialFileName:="Painel de Anotacoes.xlsm", _
              FileFilter:="Pasta de Trabalho Habilitada para Macro (*.xlsm), *.xlsm", _
              Title:="Salvar o painel como .xlsm")
    If VarType(caminho) = vbBoolean Then
        MsgBox "O arquivo ainda não foi salvo. Salve como ""Pasta de Trabalho Habilitada para Macro (*.xlsm)"" " & _
               "para não perder o código.", vbExclamation, "Instalador do Painel"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.SaveAs CStr(caminho), 52
    If Err.Number <> 0 Then
        MsgBox "Não foi possível salvar: " & Err.Description & vbCrLf & _
               "Salve manualmente como .xlsm.", vbExclamation, "Instalador do Painel"
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  Daqui para baixo: o código que o instalador grava no projeto
' ============================================================

'''
txt = head + "\n".join(gen)
# verificação: reconstruir cada componente a partir das linhas L "..."
for comp, ls in check.items():
    rec = []
    for m in re.finditer(r'^Private Sub %s_\d+\(\)\n(.*?)^End Sub' % comp, txt, re.M|re.S):
        for l in m.group(1).split('\n'):
            if l: assert l.startswith('    L "') and l.endswith('"'); rec.append(l[7:-1].replace('""','"'))
    assert rec == ls, comp
print('roundtrip ok; linhas', txt.count('\n'), 'max', max(map(len, txt.split('\n'))))
OUT.joinpath('InstalarPainel.bas').write_bytes(txt.replace('\n','\r\n').encode('cp1252'))
print("InstalarPainel.bas gerado:", OUT.joinpath("InstalarPainel.bas").stat().st_size, "bytes")
