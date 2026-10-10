# Painel de Anotações em VBA

Agenda de anotações e atas para Excel, feita em VBA com **UserForm** e **módulos de classe**.
Você escolhe o dia no calendário, registra quantas anotações quiser, marca as urgentes com um
marca-página, programa lembretes e dá baixa no que já foi resolvido, tudo num painel escuro
dentro do próprio Excel.

## O que o painel faz

- **Calendário:** clique num dia para ver e registrar as anotações daquela data. Feriados e fins de
  semana aparecem destacados, e cada dia ganha uma marca quando tem anotação, urgente, lembrete
  ativo ou tudo finalizado.
- **Anotações do dia:** cada anotação salva vira um cartão com hora, título e uma prévia do conteúdo.
  Ao salvar, o campo volta vazio para a próxima.
- **Marca-página de urgente:** a fita no canto do cartão marca ou desmarca o tema como urgente.
- **Lembrete:** duas caixas de listagem definem quando lembrar: o dia (no dia, de 1 a 7, 15 ou 30 dias
  antes) e o horário (de 15 em 15 minutos). Quando chega a hora, o cartão fica âmbar com uma
  exclamação; se a data da anotação já passou, fica vermelho.
- **Baixa:** o quadrado à esquerda do cartão finaliza a anotação: fica verde com o check, o título
  aparece riscado e o lembrete é encerrado. Clicar de novo reabre.
- **Alertas e urgentes:** lista com os lembretes ativos e os urgentes em aberto de qualquer data.
  Clicar num item leva ao dia dele.
- **Anotação aberta:** dois cliques no cartão (ou o atalho *editar*) abrem a anotação inteira para
  editar o texto, mudar o lembrete, dar baixa ou excluir.
- **Aviso ao abrir:** ao abrir o arquivo, uma mensagem lista os lembretes que já estão valendo.

Os lembretes são conferidos quando o arquivo é aberto e a cada ação no painel. Eles não disparam
com o Excel fechado.

## Instalação

Requisito: Excel para Windows.

1. Baixe o `InstalarPainel.bas` deste repositório.
2. Libere o acesso ao projeto VBA:
   **Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade →
   Configurações de Macro →** marque *Confiar no acesso ao modelo de objeto do projeto do VBA*.
3. Abra uma pasta de trabalho nova, pressione **Alt+F11** e arraste o `InstalarPainel.bas` para dentro
   do projeto (ou use **Arquivo → Importar arquivo**).
4. Clique dentro de `InstalarPainel` e pressione **F5**.
5. Escolha onde salvar o arquivo `.xlsm`. O botão **ABRIR PAINEL** já fica na aba Painel.

O instalador cria as abas **Painel**, **Base** e **Feriados**, as tabelas, os dois formulários, o módulo
e as classes. Depois de instalado, o módulo do instalador pode ser removido do projeto, e a opção do
passo 2 pode ser desmarcada.

Para atualizar para uma versão nova, importe o `InstalarPainel.bas` novo e rode `InstalarPainel` de novo.
Ele regrava o código e não mexe nas tabelas que já existem, então as anotações são mantidas.

## Como usar

- **ABRIR PAINEL** na aba Painel, ou **Alt+F8 → AbrirPainel**.
- Escreva o título e o conteúdo, escolha o lembrete se quiser e clique em **SALVAR ANOTAÇÃO**.
  A data é a do dia selecionado no calendário, e a hora é a do momento em que você salva.
- Cada anotação salva, marcada, finalizada ou excluída também salva o arquivo.
- Excluir uma anotação, ou tirar o lembrete de uma anotação que tinha, pede confirmação.
- A aba **Base** é só o banco de dados: alimente as anotações sempre pelo painel.
- Feriados estaduais, municipais e dos próximos anos podem ser acrescentados direto na tabela da aba
  **Feriados**. O instalador já traz os nacionais de 2026 e 2027.

## Estrutura do projeto

```
InstalarPainel.bas         instalador de arquivo único (traz todo o código de src/ embutido)
src/
  modPainel.bas            abre o painel, cria o botão, avisa os lembretes e preenche as caixas de dia e horário
  frmPainel.frm            painel principal: calendário, lista do dia, nova anotação, alertas e urgentes
  frmNota.frm              janela da anotação aberta com dois cliques
  clsAnotacao.cls          uma anotação e a situação do lembrete e da baixa
  clsRepositorio.cls       leitura e gravação nas tabelas tbAnotacoes e tbFeriados
  clsDiaCalendario.cls     uma casa do calendário e seus cliques
  clsCartaoNota.cls        um cartão da lista do dia: baixa, fita de urgente, editar, excluir, dois cliques
  clsItemAviso.cls         uma linha da lista de alertas e urgentes
ferramentas/
  gerar_instalador.py      monta o InstalarPainel.bas a partir de src/
```

### Como o código está organizado

Todos os controles dos formulários são criados por código quando o formulário abre, por isso os
arquivos `.frm` não têm o `.frx` que normalmente os acompanha. O instalador cria um UserForm vazio
e grava o código nele.

Controles criados assim não têm procedimentos de evento no formulário. Cada parte repetida da tela
ganha uma classe com `WithEvents`, que repassa o evento para um método público do painel:

| Classe | Evento | Chama no painel |
|---|---|---|
| `clsDiaCalendario` | clique no dia | `SelecionarDia` |
| `clsCartaoNota` | clique no quadrado de baixa | `AoAlternarFinalizada` |
| `clsCartaoNota` | clique na fita | `AoAlternarUrgente` |
| `clsCartaoNota` | dois cliques no cartão, *editar* | `AoAbrirNota` |
| `clsCartaoNota` | *excluir* | `AoExcluir` |
| `clsItemAviso` | clique no aviso | `SelecionarDia` |

Cartões e avisos são reaproveitados: o painel cria quantos precisar e esconde os que sobram, sem
nunca remover controles durante um evento.

Os formulários não mexem nas células. Toda leitura e gravação passa por `clsRepositorio`, e as regras
do lembrete e da baixa ficam em `clsAnotacao`. Bases criadas por versões anteriores ganham as colunas
novas sozinhas na primeira abertura.

### Como gerar o instalador

O `InstalarPainel.bas` é gerado a partir de `src/`. Depois de alterar um arquivo de `src/`, rode na
raiz do repositório:

```
python ferramentas/gerar_instalador.py
```

O script embute cada arquivo no instalador e confere se o texto embutido é idêntico ao original.

### Sobre a codificação dos arquivos

O `InstalarPainel.bas` e os arquivos de `src` estão em **Windows-1252 com quebras de linha CRLF**, o
formato que o editor do VBA importa sem estragar os acentos. O `.gitattributes` impede que o Git
converta esses arquivos. Se você editar o código fora do Excel, mantenha essa codificação.

## Licença

Distribuído sob a licença MIT. O texto oficial está em [LICENSE](LICENSE) e há uma tradução em [LICENCA.md](LICENCA.md).
