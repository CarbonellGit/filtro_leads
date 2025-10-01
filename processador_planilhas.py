import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os

# --- LÓGICA PRINCIPAL DE PROCESSAMENTO ---
def processar_arquivos(lista_de_arquivos, status_callback, done_callback):
    """Combina, filtra e processa uma lista de arquivos CSV.

    Esta função é projetada para ser executada em uma thread separada para não
    bloquear a interface do usuário. Ela realiza as seguintes operações:
    1.  Lê e combina múltiplos arquivos CSV em um único DataFrame do pandas.
    2.  Tenta detectar automaticamente o separador (';' ou ',').
    3.  Filtra o DataFrame combinado para remover linhas contendo e-mails de
        domínios específicos ('@colegiocarbonell.com.br', '@soucarbonell.com.br').
    4.  Remove as três primeiras colunas (A, B, C) do DataFrame resultante.
    5.  Invoca callbacks para atualizar o status na interface e para retornar
        o resultado final ou um erro.

    Args:
        lista_de_arquivos (list[str]): Uma lista de caminhos para os arquivos CSV
            a serem processados.
        status_callback (callable): Uma função de callback que aceita uma string
            para atualizar a mensagem de status na interface do usuário.
        done_callback (callable): Uma função de callback que é chamada na conclusão.
            Recebe dois argumentos: um booleano `sucesso` e o `resultado`,
            que é o DataFrame final em caso de sucesso ou uma mensagem de erro
            (string) em caso de falha.
    """
    try:
        # --- ETAPA 1: COMBINAR ARQUIVOS ---
        status_callback("Iniciando combinação dos arquivos...")
        dataframes = []
        total_arquivos = len(lista_de_arquivos)

        for i, caminho_arquivo in enumerate(lista_de_arquivos):
            status_callback(f"Lendo arquivo {i + 1}/{total_arquivos}: {os.path.basename(caminho_arquivo)}")
            try:
                # Tenta ler com separador ';' que é comum no Brasil
                df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin-1', low_memory=False)
                if df.shape[1] == 1:
                    # Se só leu uma coluna, o separador pode ser ','
                    df = pd.read_csv(caminho_arquivo, sep=',', encoding='utf-8', low_memory=False)
            except Exception:
                # Tenta com o padrão mais universal (separador ',' e encoding 'utf-8')
                df = pd.read_csv(caminho_arquivo, sep=',', encoding='utf-8', low_memory=False)
            
            dataframes.append(df)

        if not dataframes:
            raise ValueError("Nenhum arquivo CSV válido foi processado.")

        status_callback("Combinando todas as planilhas...")
        df_combinado = pd.concat(dataframes, ignore_index=True)
        status_callback(f"Planilhas combinadas! Total de {len(df_combinado)} linhas.")

        # --- ETAPA 2: FILTRAR A PLANILHA COMBINADA ---
        status_callback("Iniciando a filtragem...")

        # Identifica a coluna de email. Conforme solicitado, será a coluna E (índice 4)
        coluna_email = None
        if df_combinado.shape[1] >= 5:
            coluna_email = df_combinado.columns[4] # Índice 4 corresponde à 5ª coluna (E)
        else:
            raise ValueError("Erro: A planilha combinada não possui 5 colunas para identificar a de e-mail (Coluna E).")

        status_callback(f"Filtrando pela coluna '{coluna_email}'...")
        df_combinado[coluna_email] = df_combinado[coluna_email].astype(str)
        dominios_para_excluir = ['@colegiocarbonell.com.br', '@soucarbonell.com.br']
        condicao = df_combinado[coluna_email].str.contains('|'.join(dominios_para_excluir), na=False, case=False)
        df_filtrado = df_combinado[~condicao]
        
        linhas_removidas = len(df_combinado) - len(df_filtrado)
        status_callback(f"Filtro concluído! {linhas_removidas} linhas removidas.")
        
        # --- ETAPA 3: REMOVER COLUNAS A, B, C (NOVA MODIFICAÇÃO) ---
        status_callback("Removendo colunas desnecessárias (A, B, C)...")
        df_final = df_filtrado
        if df_filtrado.shape[1] >= 3:
            # Remove as três primeiras colunas pelos seus nomes de índice
            colunas_para_remover = df_filtrado.columns[[0, 1, 2]]
            df_final = df_filtrado.drop(columns=colunas_para_remover)
            status_callback("Colunas A, B e C removidas.")
        else:
            status_callback("Aviso: Planilha com menos de 3 colunas, não foi possível remover.")

        # Envia o DataFrame final para a função de conclusão
        done_callback(True, df_final)

    except Exception as e:
        # Envia a mensagem de erro para a função de conclusão
        done_callback(False, str(e))


# --- INTERFACE GRÁFICA (GUI) ---
class App:
    """A classe principal da aplicação de interface gráfica (GUI).

    Esta classe encapsula todos os elementos da interface do usuário (widgets)
    e a lógica para interagir com o usuário, como seleção de arquivos,
    início do processamento e exibição de status.

    Attributes:
        root (tk.Tk): A janela raiz da aplicação Tkinter.
        lista_de_arquivos (list[str]): Uma lista para armazenar os caminhos dos
            arquivos CSV selecionados pelo usuário.
        label_arquivo (ttk.Label): Widget para exibir o número de arquivos selecionados.
        btn_selecionar (ttk.Button): Botão para abrir o diálogo de seleção de arquivos.
        btn_processar (ttk.Button): Botão para iniciar o processamento dos arquivos.
        status_label (ttk.Label): Widget para exibir mensagens de status ao usuário.
    """
    def __init__(self, root):
        """Inicializa a aplicação App.

        Args:
            root (tk.Tk): A janela principal (root) da aplicação Tkinter.
        """
        self.root = root
        self.root.title("Combinador e Filtro de Leads Carbonell")
        self.root.geometry("550x300")
        self.root.resizable(False, False)

        self.lista_de_arquivos = []

        # Estilo
        style = ttk.Style(self.root)
        style.theme_use('clam')

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Label para mostrar os arquivos selecionados
        self.label_arquivo = ttk.Label(main_frame, text="Nenhum arquivo selecionado", wraplength=500, justify=tk.CENTER)
        self.label_arquivo.pack(pady=(0, 10))

        # Botão para selecionar os arquivos
        self.btn_selecionar = ttk.Button(main_frame, text="1. Selecionar Planilhas CSV", command=self.selecionar_arquivos)
        self.btn_selecionar.pack(pady=10, ipadx=10, ipady=5)

        # Botão para iniciar o processamento
        self.btn_processar = ttk.Button(main_frame, text="2. Combinar e Filtrar", state="disabled", command=self.iniciar_processamento)
        self.btn_processar.pack(pady=10, ipadx=10, ipady=5)
        
        # Label de status
        self.status_label = ttk.Label(main_frame, text="Aguardando seleção de arquivos...")
        self.status_label.pack(pady=(15, 0))

    def selecionar_arquivos(self):
        """Abre uma caixa de diálogo para o usuário selecionar múltiplos arquivos CSV.

        Atualiza a interface para mostrar o número de arquivos selecionados e
        habilita o botão de processamento.
        """
        # Abre a janela para seleção de MÚLTIPLOS arquivos CSV
        f_paths = filedialog.askopenfilenames(
            title="Selecione as planilhas CSV para combinar",
            filetypes=[("Arquivos CSV", "*.csv")]
        )
        if f_paths:
            self.lista_de_arquivos = f_paths
            self.label_arquivo.config(text=f"{len(f_paths)} arquivos selecionados.")
            self.btn_processar.config(state="normal")
            self.status_label.config(text="Arquivos prontos para processar.")

    def iniciar_processamento(self):
        """Inicia o processo de combinação e filtragem dos arquivos selecionados.

        Desabilita os botões da interface e inicia a função `processar_arquivos`
        em uma nova thread para evitar o congelamento da GUI.
        """
        if not self.lista_de_arquivos:
            messagebox.showwarning("Aviso", "Por favor, selecione os arquivos primeiro.")
            return

        self.btn_selecionar.config(state="disabled")
        self.btn_processar.config(state="disabled")
        
        # Usa threading para não congelar a interface durante o processamento
        thread = threading.Thread(
            target=processar_arquivos,
            args=(self.lista_de_arquivos, self.atualizar_status, self.processamento_concluido)
        )
        thread.start()

    def atualizar_status(self, mensagem):
        """Atualiza a label de status na interface.

        Esta função é usada como callback pela thread de processamento.

        Args:
            mensagem (str): A nova mensagem de status a ser exibida.
        """
        self.status_label.config(text=mensagem)

    def processamento_concluido(self, sucesso, resultado):
        """Manipula o resultado do processamento.

        Esta função é o callback final da thread de processamento. Se o
        processo for bem-sucedido, ela solicita ao usuário um local para
        salvar o arquivo CSV resultante. Se ocorrer um erro, exibe uma
        mensagem de erro.

        Args:
            sucesso (bool): True se o processamento foi concluído sem erros,
                False caso contrário.
            resultado (pd.DataFrame or str): O DataFrame processado em caso
                de sucesso, ou uma string com a mensagem de erro.
        """
        if sucesso:
            df_final = resultado
            # Abre a janela para o usuário escolher onde salvar
            caminho_saida = filedialog.asksaveasfilename(
                title="Salvar arquivo combinado e filtrado",
                defaultextension=".csv",
                filetypes=[("Arquivo CSV", "*.csv")],
                initialfile="planilha_combinada_e_filtrada.csv"
            )
            
            if caminho_saida:
                try:
                    # Salva o arquivo final
                    df_final.to_csv(caminho_saida, index=False, sep=';', encoding='utf-8-sig')
                    messagebox.showinfo("Sucesso!", f"Processo concluído!\n\nArquivo salvo como:\n{caminho_saida}\n\nTotal de linhas finais: {len(df_final)}")
                except Exception as e:
                    messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n\n{e}")
        else:
            messagebox.showerror("Erro no Processamento", f"Ocorreu um erro:\n\n{resultado}")
        
        # Reseta a interface para um novo uso
        self.resetar_interface()

    def resetar_interface(self):
        """Reseta a interface do usuário para o estado inicial.

        Reabilita o botão de seleção, desabilita o de processamento e limpa
        as labels de status e a lista de arquivos.
        """
        self.btn_selecionar.config(state="normal")
        self.btn_processar.config(state="disabled")
        self.label_arquivo.config(text="Nenhum arquivo selecionado")
        self.status_label.config(text="Aguardando seleção de arquivos...")
        self.lista_de_arquivos = []


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()