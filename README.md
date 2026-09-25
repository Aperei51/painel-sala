# Painel de Sala de Reuniões

Painel em tela cheia para Windows, no estilo do **Logitech Tap Scheduler**: relógio, nome da sala,
status (Disponível / Começa em X min / Em reunião), próximas reuniões da agenda e um botão
**Entrar no Teams** que abre a reunião direto no aplicativo do Teams com um toque.

![Prévia do painel](preview.png)

- Lê a agenda do **Outlook clássico** já logado (automação COM) — sem senha, sem registro no Azure
- Abre a reunião direto no app do Teams (protocolo `msteams://`), sem passar pelo navegador
- Barras laterais mudam de cor como o LED do painel físico (verde / laranja / vermelho)
- Não precisa instalar nada além do que já vem no Windows (PowerShell + WPF, arquivo único)
- Tela **Personalizar** dentro do app (engrenagem): nome da sala, subtítulo, imagem de fundo e logo

## Requisitos

- Windows 10 ou 11
- Outlook **clássico** instalado e com a conta aberta (o "novo Outlook" não permite leitura por programas)
- Microsoft Teams instalado e logado

## Instalação

1. Baixe o ZIP da versão mais recente em **Releases** e descompacte em uma pasta fixa (ex.: `C:\PainelSala`)
2. Dê dois cliques em `Instalar.cmd` → cria os atalhos na área de trabalho e a inicialização automática
3. Teste com o atalho **Painel de Sala (Demo)** (dados fictícios, sem Outlook)
4. Abra o Outlook clássico e depois o atalho **Painel de Sala**

Detalhes, configuração (`config.json`) e solução de problemas: [LEIA-ME.txt](LEIA-ME.txt).
Política de privacidade: [PRIVACIDADE.md](PRIVACIDADE.md). Publicação na Microsoft Store: [store/GUIA-STORE.md](store/GUIA-STORE.md).

## Arquivos

| Arquivo | Função |
|---|---|
| `PainelSala.ps1` | o programa (PowerShell + WPF) |
| `config.json` | configurações (nome da sala, monitor, intervalos, etc.) |
| `Instalar.cmd` / `Instalar.ps1` | cria atalhos e inicialização automática |
| `Desinstalar.cmd` | remove os atalhos |
| `PainelSala.cmd` / `Demo.cmd` | abre o painel sem instalar / modo demonstração |
| `fundo.jpg`, `logo.png`, `PainelSala.ico` | imagem de fundo, logo padrão (troque pelo seu) e ícone |

## Versões

- **1.3** — preparada para a Microsoft Store: pasta de dados gravável (`%LOCALAPPDATA%\PainelSala`) quando
  instalada pela Store, botão *Modo demonstração* nas configurações e o kit de publicação na pasta
  [`store/`](store/GUIA-STORE.md) (launcher, manifesto, ícones, script que gera o `.msix`, textos da listagem).
- **1.2** — versão neutra para distribuição: subtítulo padrão vazio, logo padrão genérico e reuniões de
  demonstração genéricas. Personalize nome, subtítulo, fundo e logo pela tela *Personalizar*.
- **1.1** — tela *Personalizar* nas configurações do app: nome da sala, subtítulo, imagem de fundo e logo,
  com prévia ao vivo e gravação automática no `config.json`. Nada mais foi alterado.
- **1.0** — versão inicial: relógio, status, próximas reuniões, Entrar no Teams, modo quiosque, instalador.
