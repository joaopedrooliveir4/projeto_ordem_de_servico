import os
import tkinter as tk
from tkinter import messagebox, ttk
from pymongo import MongoClient
from fpdf import FPDF
from pathlib import Path

# Conexão com o MongoDB
client = MongoClient('mongodb+srv://jp3066984:Hvj7kRLuRdzLuRUy@eletronicaoliveira.jn0z7.mongodb.net/?retryWrites=true&w=majority&appName=eletronicaOliveira')
db = client['eletronica_oliveira']
collection = db['ordens_servico']

# Variáveis de controle para paginação
pagina_atual = 0
itens_por_pagina = 10

# Função para buscar ordens de serviço
def buscar_ordens():
    """Buscar ordens de serviço com base na pesquisa e paginar resultados."""
    global pagina_atual
    for item in tree.get_children():
        tree.delete(item)  # Limpar a árvore de resultados

    pesquisa = entry_pesquisa.get()
    
    if pesquisa == '.':
        # Busca a última ordem de serviço salva
        ultima_ordem = collection.find().sort("numero_os", -1).limit(1)
        resultados = list(ultima_ordem)
    else:
        # Busca ordens de serviço com base na pesquisa
        resultados = list(collection.find({"$or": [
            {"numero_os": {"$regex": pesquisa, "$options": "i"}},
            {"nome": {"$regex": pesquisa, "$options": "i"}},
            {"telefone": {"$regex": pesquisa, "$options": "i"}},
            {"aparelho": {"$regex": pesquisa, "$options": "i"}}
        ]}))

    # Paginação
    total_resultados = len(resultados)
    total_paginas = (total_resultados + itens_por_pagina - 1) // itens_por_pagina  # Cálculo do total de páginas
    resultados_pagina = resultados[pagina_atual * itens_por_pagina:(pagina_atual + 1) * itens_por_pagina]

    for ordem in resultados_pagina:
        tree.insert("", "end", values=(
            ordem.get("numero_os"),
            ordem.get("nome"),
            ordem.get("endereco"),
            ordem.get("telefone"),
            ordem.get("aparelho"),
            ordem.get("servico"),
            ordem.get("pecas"),
            ordem.get("preco"),
            ordem.get("status"),
            ordem.get("aprovacao"),
            ordem.get("pagamento"),
        ))
    
    # Atualizar estado dos botões de navegação
    atualizar_botoes_paginacao(total_paginas)

def atualizar_botoes_paginacao(total_paginas):
    """Atualizar o estado dos botões de navegação de páginas."""
    btn_pagina_anterior.config(state=tk.NORMAL if pagina_atual > 0 else tk.DISABLED)
    btn_pagina_proxima.config(state=tk.NORMAL if pagina_atual < total_paginas - 1 else tk.DISABLED)

def pagina_anterior():
    """Navegar para a página anterior."""
    global pagina_atual
    if pagina_atual > 0:
        pagina_atual -= 1
        buscar_ordens()

def pagina_proxima():
    """Navegar para a próxima página."""
    global pagina_atual
    pagina_atual += 1
    buscar_ordens()

# Função para gerar PDF
def gerar_pdf():
    """Gerar um PDF com as ordens de serviço estilizado."""
    selected_item = tree.focus()  # Pega o item selecionado na tabela
    if not selected_item:
        messagebox.showerror("Erro", "Nenhuma ordem de serviço selecionada.")
        return

    valores = tree.item(selected_item, 'values')
    numero_os = valores[0]  # Pegando o "Número da OS"
    nome = valores[1]
    endereco = valores[2]
    telefone = valores[3]
    aparelho = valores[4]
    servico = valores[5]
    pecas = valores[6]
    preco = valores[7]
    status = valores[8]
    aprovacao = valores[9]
    pagamento = valores[10]

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Adicionar data atual
    from datetime import datetime
    data_atual = datetime.now().strftime("%d/%m/%Y")
    pdf.cell(0, 10, f"Data: {data_atual}", ln=True, align='R')

    # Título
    pdf.cell(0, 10, "ELETRÔNICA OLIVEIRA", ln=True, align='C')
    pdf.cell(0, 10, "MANUTENÇÃO E VENDA DE TV'S E ELETRÔNICOS", ln=True, align='C')
    pdf.cell(0, 10, "Whatsapp: (19) 98847-0072 / Telefone: (19) 3819-1253", ln=True, align='C')
    pdf.cell(0, 10, "Ordem de Serviço", ln=True, align='C')
    pdf.cell(0, 10, f"Número da OS: {numero_os}", ln=True, align='C')

    # Dados Gerais
    pdf.ln(10)  # Adiciona uma nova linha
    pdf.cell(0, 10, f"Nome: {nome}", ln=True)
    pdf.cell(0, 10, f"Endereço: {endereco}", ln=True)
    pdf.cell(0, 10, f"Telefone: {telefone}", ln=True)
    pdf.cell(0, 10, f"Aparelho: {aparelho}", ln=True)
    pdf.cell(0, 10, f"Serviço: {servico}", ln=True)
    pdf.cell(0, 10, f"Peças: {pecas}", ln=True)
    pdf.cell(0, 10, f"Preço: {preco}", ln=True)
    pdf.cell(0, 10, f"Status: {status}", ln=True)
    pdf.cell(0, 10, f"Aprovação: {aprovacao}", ln=True)
    pdf.cell(0, 10, f"Pagamento: {pagamento}", ln=True)

    # Mensagem de Garantia
    pdf.ln(10)  # Adiciona uma nova linha
    pdf.cell(0, 10, "GARANTIA DE 90 DIAS", ln=True, align='C')

    # Obter o diretório "Downloads" do usuário
    downloads_path = str(Path.home() / "Downloads")

    # Definir o caminho completo com o número da OS no nome do arquivo
    pdf_file_path = os.path.join(downloads_path, f"ordem_servico_{numero_os}.pdf")

    # Salvar o PDF
    pdf.output(pdf_file_path)

    # Exibir mensagem informando o local do arquivo
    messagebox.showinfo("Sucesso", f"PDF gerado com sucesso: {pdf_file_path}")

# Função para editar ordens de serviço
def editar_ordem():
    """Carregar os dados da ordem selecionada para edição."""    
    selected_item = tree.focus()
    if not selected_item:
        messagebox.showerror("Erro", "Nenhuma ordem de serviço selecionada.")
        return

    valores = tree.item(selected_item, 'values')

    entry_numero_os.delete(0, tk.END)
    entry_numero_os.insert(0, valores[0])

    entry_nome.delete(0, tk.END)
    entry_nome.insert(0, valores[1])

    entry_endereco.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)
    entry_telefone.insert(0, valores[3])

    entry_aparelho.delete(0, tk.END)
    entry_aparelho.insert(0, valores[4])

    entry_servico.delete(0, tk.END)
    entry_servico.insert(0, valores[5])

    entry_pecas.delete(0, tk.END)
    entry_pecas.insert(0, valores[6])

    entry_preco.delete(0, tk.END)
    entry_preco.insert(0, valores[7])

    combo_status.set(valores[8])
    combo_aprovacao.set(valores[9])
    combo_pagamento.set(valores[10])

# Função para apagar ordem de serviço
def apagar_ordem():
    """Apagar a ordem de serviço selecionada."""    
    selected_item = tree.focus()
    if not selected_item:
        messagebox.showerror("Erro", "Nenhuma ordem de serviço selecionada.")
        return

    valores = tree.item(selected_item, 'values')
    numero_os = valores[0]

    collection.delete_one({"numero_os": numero_os})
    messagebox.showinfo("Sucesso", f"Ordem de serviço {numero_os} apagada com sucesso.")

    buscar_ordens()  # Atualizar a lista após exclusão

# Função para salvar nova ordem ou atualizar uma existente
def salvar_ordem():
    """Salvar ou atualizar a ordem de serviço."""    
    numero_os = entry_numero_os.get()
    nome = entry_nome.get()
    endereco = entry_endereco.get()
    telefone = entry_telefone.get()
    aparelho = entry_aparelho.get()
    servico = entry_servico.get()
    pecas = entry_pecas.get()
    preco = entry_preco.get()
    status = combo_status.get()
    aprovacao = combo_aprovacao.get()
    pagamento = combo_pagamento.get()

    if not (numero_os and nome and telefone and aparelho and servico and status and aprovacao and pagamento):
        messagebox.showerror("Erro", "Todos os campos obrigatórios devem ser preenchidos.")
        return

    ordem_servico = {
        "numero_os": numero_os,
        "nome": nome,
        "endereco": endereco,
        "telefone": telefone,
        "aparelho": aparelho,
        "servico": servico,
        "pecas": pecas,
        "preco": preco,
        "status": status,
        "aprovacao": aprovacao,
        "pagamento": pagamento,
    }

    # Salvar nova ordem ou atualizar existente
    if numero_os.isdigit() and collection.find_one({"numero_os": numero_os}):
        collection.update_one({"numero_os": numero_os}, {"$set": ordem_servico})
        messagebox.showinfo("Sucesso", "Ordem de serviço atualizada com sucesso.")
    else:
        collection.insert_one(ordem_servico)
        messagebox.showinfo("Sucesso", "Ordem de serviço salva com sucesso.")

    # Limpar campos
    entry_numero_os.delete(0, tk.END)
    entry_nome.delete(0, tk.END)
    entry_endereco.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)
    entry_aparelho.delete(0, tk.END)
    entry_servico.delete(0, tk.END)
    entry_pecas.delete(0, tk.END)
    entry_preco.delete(0, tk.END)
    combo_status.set("")
    combo_aprovacao.set("")
    combo_pagamento.set("")

    # Atualizar lista de ordens
    buscar_ordens()

# Criação da janela principal
root = tk.Tk()
root.title("Eletrônica Oliveira")

# Frame para formulário
frame_formulario = tk.Frame(root)
frame_formulario.pack(padx=10, pady=10)

# Campos do formulário
label_numero_os = tk.Label(frame_formulario, text="Número da OS:")
label_numero_os.grid(row=0, column=0, sticky=tk.W)
entry_numero_os = tk.Entry(frame_formulario, width=30)
entry_numero_os.grid(row=0, column=1)

label_nome = tk.Label(frame_formulario, text="Nome:*")
label_nome.grid(row=1, column=0, sticky=tk.W)
entry_nome = tk.Entry(frame_formulario, width=30)
entry_nome.grid(row=1, column=1)

label_endereco = tk.Label(frame_formulario, text="Endereço:")
label_endereco.grid(row=2, column=0, sticky=tk.W)
entry_endereco = tk.Entry(frame_formulario, width=30)
entry_endereco.grid(row=2, column=1)

label_telefone = tk.Label(frame_formulario, text="Telefone:*")
label_telefone.grid(row=3, column=0, sticky=tk.W)
entry_telefone = tk.Entry(frame_formulario, width=30)
entry_telefone.grid(row=3, column=1)

label_aparelho = tk.Label(frame_formulario, text="Aparelho:*")
label_aparelho.grid(row=4, column=0, sticky=tk.W)
entry_aparelho = tk.Entry(frame_formulario, width=30)
entry_aparelho.grid(row=4, column=1)

label_servico = tk.Label(frame_formulario, text="Serviço:*")
label_servico.grid(row=5, column=0, sticky=tk.W)
entry_servico = tk.Entry(frame_formulario, width=30)
entry_servico.grid(row=5, column=1)

label_pecas = tk.Label(frame_formulario, text="Peças:")
label_pecas.grid(row=6, column=0, sticky=tk.W)
entry_pecas = tk.Entry(frame_formulario, width=30)
entry_pecas.grid(row=6, column=1)

label_preco = tk.Label(frame_formulario, text="Preço:")
label_preco.grid(row=7, column=0, sticky=tk.W)
entry_preco = tk.Entry(frame_formulario, width=30)
entry_preco.grid(row=7, column=1)

label_status = tk.Label(frame_formulario, text="Status:*")
label_status.grid(row=8, column=0, sticky=tk.W)
combo_status = ttk.Combobox(frame_formulario, values=["Orçamento a fazer", "Orçamento feito", 'Cliente retirou', 'Cliente retirou sem conserto'])
combo_status.grid(row=8, column=1)

label_aprovacao = tk.Label(frame_formulario, text="Aprovação:*")
label_aprovacao.grid(row=9, column=0, sticky=tk.W)
combo_aprovacao = ttk.Combobox(frame_formulario, values=["Pendente", "Aprovado", "Não Aprovado"])
combo_aprovacao.grid(row=9, column=1)

label_pagamento = tk.Label(frame_formulario, text="Pagamento:*")
label_pagamento.grid(row=10, column=0, sticky=tk.W)
combo_pagamento = ttk.Combobox(frame_formulario, values=["Pendente", "Pago"])
combo_pagamento.grid(row=10, column=1)

# Botões para salvar, editar, apagar e gerar PDF
btn_salvar = tk.Button(frame_formulario, text="Salvar", command=salvar_ordem)
btn_salvar.grid(row=11, column=0, padx=5, pady=5)

btn_editar = tk.Button(frame_formulario, text="Editar", command=editar_ordem)
btn_editar.grid(row=11, column=1, padx=5, pady=5)

btn_apagar = tk.Button(frame_formulario, text="Apagar", command=apagar_ordem)
btn_apagar.grid(row=11, column=2, padx=5, pady=5)

btn_pdf = tk.Button(frame_formulario, text="Gerar PDF", command=gerar_pdf)
btn_pdf.grid(row=11, column=3, padx=5, pady=5)

# Frame para pesquisa
frame_pesquisa = tk.Frame(root)
frame_pesquisa.pack(padx=10, pady=10)

label_pesquisa = tk.Label(frame_pesquisa, text="Pesquisar:")
label_pesquisa.pack(side=tk.LEFT)

entry_pesquisa = tk.Entry(frame_pesquisa, width=30)
entry_pesquisa.pack(side=tk.LEFT)

btn_buscar = tk.Button(frame_pesquisa, text="Buscar", command=buscar_ordens)
btn_buscar.pack(side=tk.LEFT)

# Tabela para exibição de ordens de serviço
tree = ttk.Treeview(root, columns=("numero_os", "nome", "endereco", "telefone", "aparelho", "servico", "pecas", "preco", "status", "aprovacao", "pagamento"), show="headings")
tree.heading("numero_os", text="Número da OS")
tree.heading("nome", text="Nome")
tree.heading("endereco", text="Endereço")
tree.heading("telefone", text="Telefone")
tree.heading("aparelho", text="Aparelho")
tree.heading("servico", text="Serviço")
tree.heading("pecas", text="Peças")
tree.heading("preco", text="Preço")
tree.heading("status", text="Status")
tree.heading("aprovacao", text="Aprovação")
tree.heading("pagamento", text="Pagamento")
tree.pack(padx=10, pady=10)

# Botões de paginação
frame_paginacao = tk.Frame(root)
frame_paginacao.pack(padx=10, pady=10)

btn_pagina_anterior = tk.Button(frame_paginacao, text="Página Anterior", command=pagina_anterior)
btn_pagina_anterior.pack(side=tk.LEFT)

btn_pagina_proxima = tk.Button(frame_paginacao, text="Próxima Página", command=pagina_proxima)
btn_pagina_proxima.pack(side=tk.LEFT)

# Iniciar o aplicativo
root.mainloop()
