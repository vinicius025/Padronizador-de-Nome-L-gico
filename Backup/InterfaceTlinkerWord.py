import tkinter as tk
from tkinter import filedialog, scrolledtext
from ScriptPadronizadorWord import aplicar_padronizacao  
from ScriptPadronizadorAuxiliar import verificar_nomes, ler_abreviacoes

def padronizar_e_verificar():
    # Limpa a caixa de texto antes de iniciar o processo
    result_box.config(state="normal")  # Habilita edição para inserir texto
    result_box.delete(1.0, tk.END)  # Limpa o conteúdo anterior
    
    # Solicita o arquivo de abreviações
    caminho_abreviacoes = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")], title="Selecione o arquivo de abreviações")
    if caminho_abreviacoes:
        # Solicita a planilha Excel
        caminho_planilha = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Selecione a planilha Excel")
        if caminho_planilha:
            # Solicita o local para salvar o arquivo Word padronizado
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")], title="Salve o arquivo Word padronizado")
            if caminho_saida:
                # Realiza a padronização do texto
                result_box.insert(tk.END, "Iniciando a padronização...\n")
                aplicar_padronizacao(caminho_abreviacoes, caminho_planilha, caminho_saida)
                result_box.insert(tk.END, "Padronização concluída e salva em um arquivo Word.\n")
                
                # Realiza a verificação de comprimento e captura as mensagens de verificação
                abreviacoes = ler_abreviacoes(caminho_abreviacoes)
                
                output_text = "Resultado da Verificação de Comprimento:\n"
                nomes_para_verificar = ["Nome_Exemplo"]  # Exemplo de lista, deve ser extraído do conteúdo real no uso final
                
                for nome in nomes_para_verificar:
                    nome_padronizado = verificar_nomes(caminho_abreviacoes, [nome])
                    if nome_padronizado:
                        output_text += nome_padronizado + "\n"
                
                # Adiciona o resultado da verificação na caixa de texto
                result_box.insert(tk.END, output_text)
                
            else:
                result_box.insert(tk.END, "Nenhum caminho de saída foi selecionado.\n")
        else:
            result_box.insert(tk.END, "Nenhuma planilha foi selecionada.\n")
    else:
        result_box.insert(tk.END, "Nenhum arquivo de abreviações foi selecionado.\n")
    
    result_box.config(state="disabled")  # Desabilita edição após inserir as mensagens

# Configurações da janela principal
root = tk.Tk()
root.title("Interface de Padronização de Dados")
root.geometry("400x300")
root.configure(bg="#f0f4f8")

# Texto de boas-vindas
label = tk.Label(root, text="Bem-vindo ao software de padronização de dados!",
                 font=("Helvetica", 12, "bold"), bg="#f0f4f8", fg="#333333")
label.pack(pady=15)

# Estilo dos botões
style_button = {
    "font": ("Helvetica", 10, "bold"),
    "bg": "#4CAF50",
    "fg": "white",
    "activebackground": "#45a049",
    "activeforeground": "white",
    "width": 25,
    "height": 2,
    "relief": "groove",
    "bd": 2
}

# Botão para padronizar e verificar o arquivo
botao_padronizacao = tk.Button(root, text="Padronizar e Verificar Arquivo", command=padronizar_e_verificar, **style_button)
botao_padronizacao.pack(pady=10)

# Caixa de texto para exibir resultados
result_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=45, height=8, font=("Helvetica", 10))
result_box.pack(pady=10)
result_box.config(state="disabled")  # Desabilita a edição inicialmente

# Rodar a interface
root.mainloop()
