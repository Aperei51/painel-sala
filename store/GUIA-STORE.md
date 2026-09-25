# Guia — publicar o Painel de Sala na Microsoft Store (pago, ~R$ 10)

O que está pronto nesta pasta `store/`:

| Arquivo | Para quê |
|---|---|
| `Empacotar-MSIX.cmd` / `.ps1` | gera o pacote `.msix` que a Store exige, com um clique |
| `store.json` | identidade do app (3 valores que o Partner Center te dá) |
| `Launcher.cs` | executável mínimo do pacote (compilado automaticamente) |
| `AppxManifest.template.xml` | manifesto do pacote (preenchido automaticamente) |
| `Assets/`, `PainelSala.ico` | ícones em todos os tamanhos exigidos |
| `listagem/` | logo da loja para a página do app |
| `LISTAGEM.md` | textos prontos para copiar na página da Store |
| `../PRIVACIDADE.md` | política de privacidade (obrigatória; já publicada no GitHub) |

O que só você pode fazer (são dados pessoais, pagamento e impostos): **passos 1, 2, 5 e 6**.
Nos passos 3, 4 e 5 eu ajudo pelo seu Chrome ou revisando o que você me enviar.

---

## 1. Criar a conta de desenvolvedor (uma vez) — ~15 min

1. Acesse **https://partner.microsoft.com/dashboard/registration** e entre com a sua conta Microsoft
   (a mesma do Windows/Outlook serve; se preferir separar, crie uma conta Microsoft nova antes).
2. Programa: **Windows & Xbox**. Tipo de conta: **Individual**.
3. País: Brasil. Nome de editor (**Publisher display name**): é o nome que aparece na loja como
   "Publicado por" — use o seu nome ou um nome fantasia (precisa ser único na Store).
4. Taxa de registro: **US$ 19 (única, sem anuidade)**, no cartão de crédito.
5. Verificação: a Microsoft confirma e-mail e telefone; contas individuais costumam ser aprovadas em minutos
   (às vezes pedem um documento — responda pelo próprio painel).

## 2. Reservar o nome do app e copiar a identidade — 5 min

1. No Partner Center: **Aplicativos e jogos > Novo produto > Aplicativo MSIX ou PWA**.
2. Nome: **Painel de Sala de Reuniões** (se estiver ocupado, tente "Painel de Sala de Reuniões — Teams e Outlook").
3. Abra **Gerenciamento do produto > Identidade do produto** e copie os três valores:
   - *Package/Identity/Name* → `IdentityName`
   - *Package/Identity/Publisher* → `Publisher` (começa com `CN=`)
   - *Package/Properties/PublisherDisplayName* → `PublisherDisplayName`
4. Cole no arquivo `store.json` (abra com o Bloco de Notas). Deixe `Version` em `1.3.0.0`.

## 3. Gerar o pacote .msix — 10 min (uma vez para instalar a ferramenta)

1. Instale as ferramentas de assinatura do Windows SDK (uma vez):
   **https://developer.microsoft.com/windows/downloads/windows-sdk/** → baixar instalador → na lista de
   componentes desmarque tudo e marque só **"Windows SDK Signing Tools for Desktop Apps"** (alguns MB).
2. Dois cliques em **`Empacotar-MSIX.cmd`**. Ele compila o launcher, monta o pacote e gera
   `store\dist\PainelSala_1.3.0.0.msix`.
3. Opcional, recomendado: **`Empacotar-MSIX.cmd -TestarLocal`** (pelo Prompt de Comando, dentro da pasta `store`)
   assina com um certificado de teste e instala o app aqui mesmo, para você ver o pacote funcionando
   como o cliente vai ver. Pede confirmação de administrador uma vez. Depois remova em Configurações > Aplicativos.

## 4. Capturas de tela — 5 min

Abra o painel (instalado ou pelo atalho), ligue **engrenagem > Modo demonstração** e tire 2 ou 3 capturas
(`Win + Shift + S`, tela inteira), 1920x1080 de preferência. Salve como PNG. Se quiser, me envie e eu reviso.

## 5. Preencher o envio no Partner Center — 20 min

Em **seu app > Iniciar envio**, siga as seções na ordem. Tudo que é texto está no `LISTAGEM.md`:

| Seção | O que fazer |
|---|---|
| **Preços e disponibilidade** | Mercados: Brasil (+ outros se quiser). Preço base: faixa mais próxima de **R$ 10,00** (o Partner Center mostra o valor em R$ ao lado de cada faixa). Sem avaliação gratuita. |
| **Propriedades** | Categoria Produtividade; política de privacidade: `https://github.com/Aperei51/painel-sala/blob/main/PRIVACIDADE.md`; site: `https://github.com/Aperei51/painel-sala`; e-mail de suporte. Hardware: toque = recomendado. |
| **Classificação etária** | Questionário IARC — responda "não" a tudo → Livre. |
| **Pacotes** | Arraste o `PainelSala_1.3.0.0.msix`. Família de dispositivos: só **Windows 10/11 Desktop**. |
| **Listagens da Store > Português (Brasil)** | Cole descrição, recursos, termos de pesquisa, requisitos; envie o logo (`listagem/LogoLoja_300x300.png`) e as capturas. |
| **Notas para certificação** | Cole o texto "Notas para a certificação" do `LISTAGEM.md` (explica o modo demonstração para o avaliador). |
| **Enviar para a Store** | Revisar e enviar. |

**Certificação:** normalmente 1 a 3 dias úteis. Se reprovar, o relatório diz o motivo — me mande que eu corrijo e gero
um pacote novo (basta subir o 3º número da versão no `store.json`, ex.: `1.3.1.0`, e rodar o `Empacotar-MSIX.cmd` de novo).

## 6. Receber os R$ 10 — configurar uma vez, antes da primeira venda

Em **Partner Center > Configurações da conta > Pagamento e imposto**:

- **Perfil de pagamento:** conta bancária no Brasil em seu nome (agência, conta, banco). Pagamento por transferência.
- **Perfil fiscal:** formulário **W-8BEN** (pessoa física fora dos EUA) — preenchido online no próprio painel;
  informe o CPF como número de identificação fiscal estrangeiro.
- **Quanto você recebe:** a Microsoft retém a comissão da Store (15% para aplicativos) e possíveis tributos;
  de R$ 10,00 ficam aproximadamente R$ 8,50 por venda. Os pagamentos são mensais, quando o saldo
  acumulado atinge o mínimo da Store (US$ 50). Detalhes atualizados em
  https://learn.microsoft.com/partner-center/payout-statement.
- A nota fiscal ao comprador é emitida pela Microsoft (ela é a loja); você não emite NF por venda.

---

## Depois de publicado

- **Atualizar o app:** eu gero a nova versão (ex.: 1.4.0.0), você roda `Empacotar-MSIX.cmd` e envia o `.msix` em
  *Atualizar* no Partner Center. A Store atualiza os clientes automaticamente.
- **Distribuição fora da Store** continua funcionando com o ZIP do GitHub (o `Instalar.cmd`), sem custo.

## Se algo travar

| Sintoma | Solução |
|---|---|
| `store.json` "PREENCHA" | copie os 3 valores da Identidade do produto (passo 2) |
| "MakeAppx.exe não encontrado" | instale o componente do SDK (passo 3.1) e rode de novo |
| Partner Center rejeita o pacote por identidade | `IdentityName`/`Publisher` diferentes dos reservados — copie de novo, exatamente |
| Reprovado na certificação | me envie o relatório; quase sempre é ajuste de listagem ou nota para o avaliador |
