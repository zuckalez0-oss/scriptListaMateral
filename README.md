# Analisador de Lista de Materiais

Este aplicativo automatiza o processo de análise e transferência de dados de listas de materiais, extraídas do software MCalc (em formato `.docx`), para uma planilha Excel padronizada. A ferramenta foi projetada para otimizar o tempo e reduzir erros manuais no processo de orçamentação e controle de materiais.

## Funcionalidades Principais

-   **Interface Gráfica Intuitiva:** Uma interface simples construída com Tkinter que permite ao usuário selecionar facilmente os arquivos de entrada e acompanhar o progresso da automação através de um log de eventos.

-   **Extração de Dados de Arquivos Word (.docx):**
    -   Lê automaticamente a tabela de materiais gerada pelo MCalc.
    -   Extrai informações essenciais como descrição do perfil, tipo de aço, comprimento e peso.

-   **Classificação e Mapeamento Inteligente de Perfis:**
    -   Identifica e categoriza diferentes tipos de perfis de aço (Viga W, Perfil U, Terça, Cantoneira, Tubo) com base em suas descrições.
    -   Normaliza nomenclaturas, como a de Vigas W (ex: "W 150 x 22.5" para "W150X22,5"), para garantir a correspondência correta na planilha de destino.

-   **Preenchimento Automático da Planilha Excel:**
    -   Carrega uma planilha Excel modelo.
    -   Localiza a linha correta para cada item da lista de materiais:
        -   Para **Vigas W**, a busca é feita pelo nome normalizado do perfil.
        -   Para **outros perfis**, o sistema encontra a próxima linha vazia na seção correspondente.
    -   Insere os dados extraídos e processados nas colunas apropriadas (dimensões, espessura, tipo de aço, comprimento, peso).

-   **Processamento de Dados e Regras de Negócio:**
    -   Extrai dimensões (abas, alma, espessura) da descrição do perfil.
    -   Aplica regras específicas, como a duplicação do comprimento para perfis do tipo "CA".
    -   Possui uma lógica híbrida para determinar o comprimento, priorizando a coluna de comprimento total, mas sendo capaz de extraí-lo da descrição do perfil se necessário (ex: "C=5000").

-   **Geração de Relatório Final:**
    -   Salva uma nova planilha com o sufixo `_processado` no nome, preservando o arquivo original.
    -   Oculta automaticamente as linhas que não foram preenchidas na planilha final, gerando um relatório limpo e de fácil visualização.

## Como Usar

1.  Execute o aplicativo.
2.  Na seção "Lista de Material (.docx)", clique em **Procurar...** e selecione o arquivo Word exportado pelo MCalc.
3.  Na seção "Planilha Excel (.xlsx)", clique em **Procurar...** e selecione o arquivo Excel que servirá como modelo.
4.  Clique no botão **Iniciar Script**.
5.  Aguarde a mensagem de sucesso. Uma nova planilha com o nome `[nome_original]_processado.xlsx` será criada no mesmo diretório da planilha modelo.

## Tecnologias Utilizadas

-   **Python 3**
-   **Tkinter:** Para a interface gráfica.
-   **python-docx:** Para a leitura e extração de dados de arquivos `.docx`.
-   **openpyxl:** Para a manipulação e escrita em arquivos `.xlsx`.
