# Combinador e Filtro de Planilhas de Leads

## Propósito

Esta é uma ferramenta de automação desenvolvida em Python com uma interface gráfica (GUI) usando Tkinter. O objetivo principal é ajudar a equipe de marketing e vendas a processar e limpar listas de leads exportadas de diversas fontes no formato CSV.

A aplicação realiza as seguintes tarefas:
1.  **Combina Múltiplos Arquivos**: Permite ao usuário selecionar vários arquivos CSV, que são combinados em uma única planilha.
2.  **Filtra Leads Internos**: Remove automaticamente contatos que possuem e-mails de domínios internos (ex: `@colegiocarbonell.com.br`, `@soucarbonell.com.br`), garantindo que a lista final contenha apenas leads externos.
3.  **Remove Colunas Irrelevantes**: Exclui as três primeiras colunas (A, B, C) das planilhas, que geralmente contêm informações não essenciais para a prospecção.
4.  **Exporta o Resultado**: Salva a planilha limpa e processada em um novo arquivo CSV, pronta para ser utilizada em campanhas de marketing ou outras análises.

## Setup

Para executar esta aplicação, você precisará ter o Python instalado em seu computador, bem como a biblioteca `pandas`.

### Pré-requisitos

-   **Python 3**: Certifique-se de que o Python 3 está instalado. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).

### Instalação de Dependências

A única dependência externa é a biblioteca `pandas`. Você pode instalá-la usando o gerenciador de pacotes `pip`. Abra seu terminal ou prompt de comando e execute o seguinte comando:

```bash
pip install pandas
```

## Como Usar

A aplicação possui uma interface gráfica simples e intuitiva. Para utilizá-la, siga os passos abaixo:

1.  **Execute o Script**: Abra um terminal ou prompt de comando, navegue até o diretório onde o arquivo `processador_planilhas.py` está localizado e execute o seguinte comando:

    ```bash
    python processador_planilhas.py
    ```

2.  **Selecione as Planilhas**:
    -   Clique no botão **"1. Selecionar Planilhas CSV"**.
    -   Uma janela de diálogo será aberta. Navegue até a pasta onde seus arquivos de leads estão, selecione um ou mais arquivos CSV e clique em "Abrir".
    -   A interface informará quantos arquivos foram selecionados.

3.  **Inicie o Processamento**:
    -   Com os arquivos selecionados, o botão **"2. Combinar e Filtrar"** será habilitado. Clique nele para iniciar o processo.
    -   A aplicação mostrará o status do processamento em tempo real (lendo, combinando, filtrando, etc.).

4.  **Salve o Arquivo Final**:
    -   Após a conclusão do processamento, uma janela de diálogo para salvar o arquivo será exibida.
    -   Escolha o local onde deseja salvar a planilha processada, defina um nome para o arquivo (o padrão é `planilha_combinada_e_filtrada.csv`) e clique em "Salvar".

5.  **Conclusão**:
    -   Uma mensagem de sucesso aparecerá, confirmando que o arquivo foi salvo e informando o número total de linhas na planilha final.
    -   A interface será resetada, pronta para um novo uso.

Em caso de erro durante o processo, uma mensagem de alerta será exibida com os detalhes do problema.