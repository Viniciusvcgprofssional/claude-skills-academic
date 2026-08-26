---
name: word-personalize
description: >
  Cria, gera e gerencia modelos personalizados de documentos Word (.docx) prontos para salvar como
    templates (.dotx) no Microsoft Word. Use esta skill SEMPRE que o usuário pedir para criar, gerar,
      atualizar ou personalizar modelos de documentos Word — incluindo modelos ABNT (TCC, artigo
        científico, carta, ofício, relatório, projeto de pesquisa, requerimento), modelos institucionais
          (universidades e outras instituições), ou qualquer documento Word customizado com formatação específica.
            Também ativa quando o usuário mencionar ".dotx", "modelo do Word", "template Word", "modelo
              personalizado", "documento ABNT", "formatar TCC", "modelo de carta", "normas ABNT Word" ou
                variações dessas expressões. Gera os arquivos .docx diretamente para download.
                ---

                # Word Personalize

                Skill para criação de modelos Word personalizados com formatação profissional.

                ## Visão Geral

                Esta skill gera arquivos `.docx` prontos para serem salvos como modelos `.dotx` no Word.
                Cada modelo é criado via Node.js com a biblioteca `docx`, seguindo as normas ABNT ou as
                especificações institucionais solicitadas.

                ## Fluxo de Trabalho

                1. **Identificar o tipo de documento** solicitado
                2. **Verificar se há especificação institucional** (ex: uma universidade específica) — se sim, pesquisar na web as normas específicas
                3. **Gerar o arquivo .docx** com o script Node.js
                4. **Apresentar o arquivo** para download com instruções de instalação como template

                ## Padrões ABNT Aplicados em Todos os Modelos

                | Elemento | Valor |
                |---|---|
                | Fonte | Times New Roman 12pt (padrão) ou Arial 12pt |
                | Margem superior e esquerda | 3 cm (1701 DXA) |
                | Margem inferior e direita | 2 cm (1134 DXA) |
                | Espaçamento entre linhas | 1,5 (exceto resumo, referências e citações longas = simples) |
                | Recuo de parágrafo | 1,25 cm (709 DXA) |
                | Alinhamento do texto | Justificado |
                | Numeração de página | Algarismos arábicos, centralizado no rodapé |
                | Tamanho do papel | A4 (11906 x 16838 DXA) |

                ## Tipos de Documento e Normas

                ### TCC / Monografia (NBR 14724)
                **Elementos obrigatórios (pré-textuais):** Capa, folha de rosto*, resumo PT + EN, sumário
                **Elementos obrigatórios (textuais):** Introdução, desenvolvimento, conclusão
                **Elementos obrigatórios (pós-textuais):** Referências
                **Opcionais:** Errata, dedicatória, agradecimentos, epígrafe, listas, apêndices, anexos
                *Verso da folha de rosto = ficha catalográfica (elaborada pela biblioteca)

                ### Artigo Científico (NBR 6022)
                Título, autores + afiliação, resumo/abstract (simples, 100–250 palavras), palavras-chave,
                seções numeradas, referências. Sem numeração nas páginas pré-textuais.

                ### Relatório Técnico (NBR 10719)
                Capa com identificação, objetivo, contextualização, metodologia, resultados, conclusão,
                referências. Cabeçalho com número do relatório.

                ### Projeto de Pesquisa (NBR 15287)
                Tema, problema, hipótese, objetivos (geral + específicos), justificativa, referencial teórico,
                metodologia, cronograma, orçamento (se aplicável), referências.

                ### Carta / Ofício
                Timbre, local e data, assunto, destinatário, vocativo, corpo (3 parágrafos mínimo), fecho,
                assinatura. Espaçamento 1,5, sem recuo nos parágrafos da carta.

                ### Requerimento
                Qualificação completa do requerente, objeto do pedido, fecho ("Nestes termos, pede deferimento"),
                local/data, assinatura.

                ## Exemplo de Especificação Institucional — UFAL (SIBI/UFAL 2022)

                A UFAL segue a ABNT padrão com os seguintes **acréscimos obrigatórios**:

                - **Ficha catalográfica**: obrigatória no verso da folha de rosto, gerada pelo SIBI/UFAL (não pelo aluno)
                - **Folha de aprovação**: deve estar assinada por pelo menos 2 membros da banca ou 1 membro + coordenador
                - **Lombada** (quando impresso em capa dura): nome do autor, título e ano longitudinalmente
                - **Depósito no RIUFAL**: após defesa, enviar para `ri@sibi.ufal.br` com Termo de Autorização assinado
                - Modelos editáveis oficiais disponíveis em: `sibi.ufal.br/portal/?page_id=1770`
                - O SIBI oferece minicurso semestral gratuito de normalização ("Desvendando o Manual UFAL")

                ## Como Gerar um Modelo

                ### Setup (verificar antes de rodar)
                ```bash
                npm list -g docx 2>/dev/null | head -3 || npm install -g docx
                ```

                ### Template do Script Node.js
                ```javascript
                const {
                  Document, Packer, Paragraph, TextRun, Header, Footer,
                    AlignmentType, HeadingLevel, NumberFormat, PageBreak, PageNumberElement
                    } = require('docx');
                    const fs = require('fs');

                    // Constantes ABNT
                    const A4_W = 11906, A4_H = 16838;
                    const MARGIN = { top: 1701, right: 1134, bottom: 1134, left: 1701 };
                    const FONT = 'Times New Roman';
                    const SIZE_NORMAL = 24, SIZE_TITULO = 28, SIZE_SMALL = 20;
                    const SPACING_15 = { line: 360, lineRule: 'auto', before: 0, after: 0 };

                    // Helpers
                    const run = (text, opts = {}) => new TextRun({ text, font: FONT, size: SIZE_NORMAL, ...opts });
                    const campo = (label) => new TextRun({ text: `[${label}]`, font: FONT, size: SIZE_NORMAL, color: '0070C0', underline: {} });
                    const paragrafo = (children, opts = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], spacing: SPACING_15, ...opts });
                    const linhaBranca = () => new Paragraph({ children: [new TextRun('')], spacing: { line: 360 } });
                    ```

                    **Regras críticas da biblioteca `docx`:**
                    - Nunca usar `\n` — usar `Paragraph` separados
                    - Usar `PageNumberElement()` para numeração de página (não `PageNumber`)
                    - Tabelas sempre com `WidthType.DXA`, nunca `PERCENTAGE`
                    - Campos editáveis: cor `0070C0` (azul) com underline
                    - Após gerar, copiar para `/mnt/user-data/outputs/`

                    ## Como Instalar no Word (instruir o usuário)

                    1. Baixar o arquivo `.docx`
                    2. Abrir no Word
                    3. **Arquivo → Salvar como → Tipo: Modelo do Word (*.dotx)**
                    4. Aceitar a pasta padrão sugerida pelo Word
                    5. Usar via **Arquivo → Novo → Modelos pessoais**

                    **Pasta manual (Windows):** `C:\Users\[Nome]\AppData\Roaming\Microsoft\Templates`

                    ## Quando Pesquisar Normas Institucionais

                    Se o usuário mencionar uma universidade ou instituição específica:
                    1. Buscar no site oficial da biblioteca/SIBI da instituição
                    2. Procurar por "manual de normalização", "modelo TCC" ou "normas ABNT [instituição]"
                    3. Identificar exigências extras além da ABNT padrão
                    4. Incorporar essas exigências no modelo gerado

                    ## Campos Azuis nos Modelos

                    Todos os modelos usam campos em **azul sublinhado** `[Entre colchetes]` para indicar onde o
                    usuário deve preencher. Isso torna o modelo intuitivo — basta substituir cada campo azul pelo
                    conteúdo real.
                    
