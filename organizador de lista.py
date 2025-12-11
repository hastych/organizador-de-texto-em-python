import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import re
import pandas as pd
from datetime import datetime

class CandidatoOrganizador:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de Candidatos - IFNMG Processo 1048/2025")
        self.root.geometry("1300x850")
        
        self.candidatos = []
        self.df_candidatos = None
        self.arquivo_carregado = None
        
        # Configurar tema
        self.setup_theme()
        self.setup_interface()
        
    def setup_theme(self):
        """Configura cores e estilo da interface"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Cores personalizadas
        self.cor_deferido = '#e6ffe6'
        self.cor_indeferido = '#ffe6e6'
        self.cor_destaque = '#e6f3ff'
        
    def setup_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(header_frame, text="📋 ORGANIZADOR DE CANDIDATOS", 
                 font=('Arial', 16, 'bold')).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Label(header_frame, text="IFNMG - Processo Seletivo 1048/2025", 
                 font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W)
        
        # Frame de controles principais
        control_frame = ttk.LabelFrame(main_frame, text="Controles Principais", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Linha 1 de controles
        ttk.Button(control_frame, text="📂 CARREGAR ARQUIVO TXT", 
                  command=self.carregar_arquivo, width=25).grid(row=0, column=0, padx=5, pady=5)
        
        self.arquivo_label = ttk.Label(control_frame, text="Nenhum arquivo carregado", 
                                      foreground="gray")
        self.arquivo_label.grid(row=0, column=1, padx=10, sticky=tk.W)
        
        # Linha 2 de controles - Ordenação
        ttk.Label(control_frame, text="Ordenar por:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.ordenar_var = tk.StringVar(value="nota_desc")
        ordenar_combo = ttk.Combobox(control_frame, textvariable=self.ordenar_var,
                                    values=["Nota (Mais Alta)", "Nota (Mais Baixa)", 
                                            "Nome (A-Z)", "Nome (Z-A)", 
                                            "Inscrição", "Situação"],
                                    width=20, state="readonly")
        ordenar_combo.grid(row=1, column=1, padx=5, pady=5)
        ordenar_combo.bind('<<ComboboxSelected>>', self.aplicar_ordenacao)
        
        # Linha 3 de controles - Filtros
        ttk.Label(control_frame, text="Filtrar:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Filtro de situação
        self.filtro_situacao_var = tk.StringVar(value="TODOS")
        situacao_combo = ttk.Combobox(control_frame, textvariable=self.filtro_situacao_var,
                                     values=["TODOS", "DEFERIDA", "INDEFERIDA"], 
                                     width=15, state="readonly")
        situacao_combo.grid(row=2, column=1, padx=5, pady=5)
        situacao_combo.bind('<<ComboboxSelected>>', self.aplicar_filtros)
        
        # Filtro de modalidade
        ttk.Label(control_frame, text="Modalidade:").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        self.filtro_modalidade_var = tk.StringVar(value="TODAS")
        self.modalidade_combo = ttk.Combobox(control_frame, textvariable=self.filtro_modalidade_var,
                                           width=20, state="readonly")
        self.modalidade_combo.grid(row=2, column=3, padx=5, pady=5)
        self.modalidade_combo.bind('<<ComboboxSelected>>', self.aplicar_filtros)
        
        # Busca
        ttk.Label(control_frame, text="Buscar:").grid(row=2, column=4, padx=5, pady=5, sticky=tk.W)
        self.busca_var = tk.StringVar()
        busca_entry = ttk.Entry(control_frame, textvariable=self.busca_var, width=25)
        busca_entry.grid(row=2, column=5, padx=5, pady=5)
        ttk.Button(control_frame, text="🔍", 
                  command=self.aplicar_filtros, width=3).grid(row=2, column=6, padx=5, pady=5)
        
        # Botões de ação
        action_frame = ttk.Frame(control_frame)
        action_frame.grid(row=3, column=0, columnspan=7, pady=10)
        
        ttk.Button(action_frame, text="📊 VER ESTATÍSTICAS", 
                  command=self.mostrar_estatisticas).grid(row=0, column=0, padx=5)
        
        ttk.Button(action_frame, text="💾 EXPORTAR CSV", 
                  command=self.exportar_csv).grid(row=0, column=1, padx=5)
        
        ttk.Button(action_frame, text="📄 EXPORTAR TXT", 
                  command=self.exportar_txt).grid(row=0, column=2, padx=5)
        
        ttk.Button(action_frame, text="🖨️ IMPRIMIR LISTA", 
                  command=self.imprimir_lista).grid(row=0, column=3, padx=5)
        
        ttk.Button(action_frame, text="🔄 LIMPAR FILTROS", 
                  command=self.limpar_filtros).grid(row=0, column=4, padx=5)
        
        # Frame da tabela
        table_frame = ttk.LabelFrame(main_frame, text="Lista de Candidatos", padding="5")
        table_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")
        
        # Treeview
        self.tree = ttk.Treeview(table_frame, 
                                columns=("Inscrição", "CPF", "Nome", "Modalidade", "Nota", "Situação", "Motivo"),
                                yscrollcommand=vsb.set,
                                xscrollcommand=hsb.set,
                                selectmode="extended",
                                height=20)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configurar colunas
        columns = {
            "#0": {"text": "#", "width": 50, "minwidth": 40, "anchor": tk.CENTER},
            "Inscrição": {"text": "INSCRIÇÃO", "width": 90, "minwidth": 80, "anchor": tk.CENTER},
            "CPF": {"text": "CPF", "width": 140, "minwidth": 120, "anchor": tk.CENTER},
            "Nome": {"text": "NOME DO CANDIDATO", "width": 300, "minwidth": 200, "anchor": tk.W},
            "Modalidade": {"text": "MODALIDADE", "width": 150, "minwidth": 120, "anchor": tk.CENTER},
            "Nota": {"text": "NOTA", "width": 80, "minwidth": 70, "anchor": tk.CENTER},
            "Situação": {"text": "SITUAÇÃO", "width": 100, "minwidth": 90, "anchor": tk.CENTER},
            "Motivo": {"text": "MOTIVO / OBSERVAÇÃO", "width": 350, "minwidth": 250, "anchor": tk.W}
        }
        
        for col, config in columns.items():
            if col == "#0":
                self.tree.heading(col, text=config["text"])
                self.tree.column(col, width=config["width"], minwidth=config["minwidth"], anchor=config["anchor"])
            else:
                self.tree.heading(col, text=config["text"])
                self.tree.column(col, width=config["width"], minwidth=config["minwidth"], anchor=config["anchor"])
        
        # Grid da treeview
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Frame de informações
        info_frame = ttk.Frame(main_frame)
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        # Contadores
        self.total_label = ttk.Label(info_frame, text="Total: 0", font=('Arial', 10, 'bold'))
        self.total_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        
        self.deferidos_label = ttk.Label(info_frame, text="Deferidos: 0", foreground="green", font=('Arial', 10))
        self.deferidos_label.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)
        
        self.indeferidos_label = ttk.Label(info_frame, text="Indeferidos: 0", foreground="red", font=('Arial', 10))
        self.indeferidos_label.grid(row=0, column=2, padx=10, pady=5, sticky=tk.W)
        
        self.nota_max_label = ttk.Label(info_frame, text="Maior nota: 0.00", font=('Arial', 10))
        self.nota_max_label.grid(row=0, column=3, padx=10, pady=5, sticky=tk.W)
        
        self.nota_min_label = ttk.Label(info_frame, text="Menor nota (deferidos): 0.00", font=('Arial', 10))
        self.nota_min_label.grid(row=0, column=4, padx=10, pady=5, sticky=tk.W)
        
        # Status bar
        self.status_var = tk.StringVar(value="Pronto para carregar arquivo")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, padding=5)
        status_bar.grid(row=4, column=0, sticky=(tk.W, tk.E))
        
        # Configurar expansão
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
    def carregar_arquivo(self):
        """Carrega e processa o arquivo TXT"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo de candidatos",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            try:
                with open(arquivo, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                
                self.arquivo_carregado = arquivo
                self.arquivo_label.config(text=f"📄 {arquivo.split('/')[-1]}")
                
                # Processar o arquivo
                self.candidatos = self.processar_arquivo(conteudo)
                self.df_candidatos = pd.DataFrame(self.candidatos)
                
                # Atualizar interface
                self.atualizar_treeview()
                
                # Atualizar combobox de modalidades
                modalidades = ["TODAS"] + sorted(list(set([c['modalidade'] for c in self.candidatos if c['modalidade']])))
                self.modalidade_combo['values'] = modalidades
                
                # Atualizar estatísticas
                self.atualizar_contadores()
                
                self.status_var.set(f"✅ Arquivo carregado: {len(self.candidatos)} candidatos processados")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar arquivo:\n{str(e)}")
    
    def processar_arquivo(self, conteudo):
        """Processa o conteúdo do arquivo TXT"""
        candidatos = []
        linhas = conteudo.split('\n')
        
        i = 0
        total_linhas = len(linhas)
        
        while i < total_linhas:
            linha = linhas[i].strip()
            
            # Padrão para identificar linha de candidato: número_inscrição + CPF mascarado
            # Exemplo: "10678 ###.###.528-42 ABEL OLIVEIRA DA SILVA"
            if re.match(r'^\d{4,6}\s+###\.###\.\d{3}-\d{2}', linha):
                try:
                    # Dividir a linha
                    partes = linha.split(maxsplit=2)
                    if len(partes) >= 3:
                        inscricao = partes[0]
                        cpf = partes[1]
                        nome = partes[2]
                        
                        # Verificar continuação do nome
                        j = i + 1
                        while j < total_linhas and linhas[j].strip() and not re.match(r'^\d{4,6}\s+###\.###\.\d{3}-\d{2}', linhas[j]):
                            nome += " " + linhas[j].strip()
                            j += 1
                        
                        nome = nome.strip()
                        
                        # Procurar dados acadêmicos
                        modalidade = ""
                        nota = 0.0
                        situacao = ""
                        motivo = ""
                        
                        # Procurar nas próximas 10 linhas
                        for k in range(i, min(i + 10, total_linhas)):
                            linha_k = linhas[k].strip()
                            
                            # Verificar modalidade
                            modalidades_possiveis = [
                                'Ampla Concorrência', 'LI_EP', 'LB_EP', 'LI_PPI', 'LB_PPI',
                                'LI_PCD', 'LB_PCD', 'V_PCD', 'V_EFA', 'LB_Q'
                            ]
                            
                            for mod in modalidades_possiveis:
                                if mod in linha_k:
                                    modalidade = mod
                                    break
                            
                            # Verificar situação e nota
                            if 'DEFERIDA' in linha_k or 'INDEFERIDA' in linha_k:
                                if 'DEFERIDA' in linha_k:
                                    situacao = 'DEFERIDA'
                                    # Padrão: "Ampla Concorrência 89.00 DEFERIDA -"
                                    padrao_nota = r'(\d+\.\d{2}|\d+)\s+DEFERIDA'
                                else:
                                    situacao = 'INDEFERIDA'
                                    # Padrão: "Ampla Concorrência 0.00 INDEFERIDA MOTIVO"
                                    padrao_nota = r'(\d+\.\d{2}|\d+)\s+INDEFERIDA'
                                
                                # Extrair nota
                                match_nota = re.search(padrao_nota, linha_k)
                                if match_nota:
                                    try:
                                        nota = float(match_nota.group(1))
                                    except:
                                        nota = 0.0
                                
                                # Extrair motivo para indeferidos
                                if situacao == 'INDEFERIDA':
                                    partes_ind = linha_k.split('INDEFERIDA', 1)
                                    if len(partes_ind) > 1:
                                        motivo = partes_ind[1].strip(' -')
                                    
                                    # Verificar próxima linha para motivo completo
                                    if k + 1 < total_linhas and linhas[k+1].strip():
                                        prox_linha = linhas[k+1].strip()
                                        if not re.match(r'^\d{4,6}\s+###\.###\.\d{3}-\d{2}', prox_linha):
                                            motivo += " " + prox_linha
                                
                                break  # Encontrou situação
                        
                        # Para deferidos sem motivo, colocar "-"
                        if situacao == 'DEFERIDA':
                            motivo = "-"
                        
                        # Adicionar candidato
                        candidato = {
                            'inscricao': inscricao,
                            'cpf': cpf,
                            'nome': nome,
                            'modalidade': modalidade,
                            'nota': nota,
                            'situacao': situacao,
                            'motivo': motivo.strip()
                        }
                        
                        candidatos.append(candidato)
                        
                except Exception as e:
                    print(f"Erro ao processar linha {i}: {e}")
                    print(f"Conteúdo: {linha}")
            
            i += 1
        
        return candidatos
    
    def atualizar_treeview(self, candidatos_filtrados=None):
        """Atualiza a Treeview com os candidatos"""
        # Limpar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Determinar quais candidatos mostrar
        if candidatos_filtrados is None:
            candidatos_a_mostrar = self.candidatos
        else:
            candidatos_a_mostrar = candidatos_filtrados
        
        # Aplicar ordenação atual
        ordenacao = self.ordenar_var.get()
        if ordenacao == "Nota (Mais Alta)":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: x['nota'], reverse=True)
        elif ordenacao == "Nota (Mais Baixa)":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: x['nota'])
        elif ordenacao == "Nome (A-Z)":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: x['nome'])
        elif ordenacao == "Nome (Z-A)":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: x['nome'], reverse=True)
        elif ordenacao == "Inscrição":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: int(x['inscricao']))
        elif ordenacao == "Situação":
            candidatos_a_mostrar = sorted(candidatos_a_mostrar, key=lambda x: x['situacao'])
        
        # Adicionar itens à treeview
        for idx, cand in enumerate(candidatos_a_mostrar, 1):
            nota_str = f"{cand['nota']:.2f}" if cand['nota'] > 0 else "0.00"
            
            # Tag para colorir
            tag = 'deferido' if cand['situacao'] == 'DEFERIDA' else 'indeferido'
            
            self.tree.insert("", "end", text=str(idx),
                           values=(cand['inscricao'],
                                   cand['cpf'],
                                   cand['nome'],
                                   cand['modalidade'],
                                   nota_str,
                                   cand['situacao'],
                                   cand['motivo']),
                           tags=(tag,))
        
        # Configurar cores das tags
        self.tree.tag_configure('deferido', background=self.cor_deferido)
        self.tree.tag_configure('indeferido', background=self.cor_indeferido)
    
    def atualizar_contadores(self):
        """Atualiza os contadores de estatísticas"""
        if not self.candidatos:
            return
        
        total = len(self.candidatos)
        deferidos = sum(1 for c in self.candidatos if c['situacao'] == 'DEFERIDA')
        indeferidos = sum(1 for c in self.candidatos if c['situacao'] == 'INDEFERIDA')
        
        # Notas dos deferidos
        notas_deferidos = [c['nota'] for c in self.candidatos if c['situacao'] == 'DEFERIDA' and c['nota'] > 0]
        nota_max = max(notas_deferidos) if notas_deferidos else 0
        nota_min = min(notas_deferidos) if notas_deferidos else 0
        
        # Atualizar labels
        self.total_label.config(text=f"Total: {total}")
        self.deferidos_label.config(text=f"Deferidos: {deferidos}")
        self.indeferidos_label.config(text=f"Indeferidos: {indeferidos}")
        self.nota_max_label.config(text=f"Maior nota: {nota_max:.2f}")
        self.nota_min_label.config(text=f"Menor nota (deferidos): {nota_min:.2f}")
    
    def aplicar_ordenacao(self, event=None):
        """Aplica a ordenação selecionada"""
        self.atualizar_treeview()
    
    def aplicar_filtros(self, event=None):
        """Aplica todos os filtros selecionados"""
        if not self.candidatos:
            return
        
        candidatos_filtrados = self.candidatos.copy()
        
        # Filtrar por situação
        situacao = self.filtro_situacao_var.get()
        if situacao != "TODOS":
            candidatos_filtrados = [c for c in candidatos_filtrados if c['situacao'] == situacao]
        
        # Filtrar por modalidade
        modalidade = self.filtro_modalidade_var.get()
        if modalidade != "TODAS":
            candidatos_filtrados = [c for c in candidatos_filtrados if c['modalidade'] == modalidade]
        
        # Filtrar por busca
        busca = self.busca_var.get().strip().upper()
        if busca:
            candidatos_filtrados = [c for c in candidatos_filtrados if busca in c['nome'].upper() or 
                                   busca in c['inscricao'] or 
                                   busca in c['cpf']]
        
        # Atualizar treeview com filtros
        self.atualizar_treeview(candidatos_filtrados)
        self.status_var.set(f"Mostrando {len(candidatos_filtrados)} candidatos (filtrados)")
    
    def limpar_filtros(self):
        """Limpa todos os filtros"""
        self.filtro_situacao_var.set("TODOS")
        self.filtro_modalidade_var.set("TODAS")
        self.busca_var.set("")
        self.ordenar_var.set("nota_desc")
        self.atualizar_treeview()
        self.status_var.set("Filtros limpos")
    
    def mostrar_estatisticas(self):
        """Mostra estatísticas detalhadas"""
        if not self.candidatos:
            messagebox.showinfo("Estatísticas", "Nenhum dado carregado.")
            return
        
        total = len(self.candidatos)
        deferidos = sum(1 for c in self.candidatos if c['situacao'] == 'DEFERIDA')
        indeferidos = total - deferidos
        
        # Notas
        notas_deferidos = [c['nota'] for c in self.candidatos if c['situacao'] == 'DEFERIDA']
        nota_max = max(notas_deferidos) if notas_deferidos else 0
        nota_min = min(notas_deferidos) if notas_deferidos else 0
        nota_media = sum(notas_deferidos) / len(notas_deferidos) if notas_deferidos else 0
        
        # Modalidades
        modalidades = {}
        for c in self.candidatos:
            if c['modalidade']:
                modalidades[c['modalidade']] = modalidades.get(c['modalidade'], 0) + 1
        
        # Construir mensagem
        msg = f"""📊 ESTATÍSTICAS DETALHADAS
        
        Total de Candidatos: {total}
        • Deferidos: {deferidos} ({deferidos/total*100:.1f}%)
        • Indeferidos: {indeferidos} ({indeferidos/total*100:.1f}%)
        
        Notas (Deferidos):
        • Maior: {nota_max:.2f}
        • Menor: {nota_min:.2f}
        • Média: {nota_media:.2f}
        
        Distribuição por Modalidade:
        """
        
        for mod, qtd in sorted(modalidades.items()):
            msg += f"• {mod}: {qtd} candidatos\n"
        
        messagebox.showinfo("Estatísticas Detalhadas", msg)
    
    def exportar_csv(self):
        """Exporta os dados para CSV"""
        if not self.candidatos:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar!")
            return
        
        arquivo = filedialog.asksaveasfilename(
            title="Salvar como CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            try:
                # Obter dados atuais da treeview
                dados = []
                for item in self.tree.get_children():
                    valores = self.tree.item(item)['values']
                    dados.append(valores)
                
                df = pd.DataFrame(dados, columns=["Inscrição", "CPF", "Nome", "Modalidade", "Nota", "Situação", "Motivo"])
                df.to_csv(arquivo, index=False, encoding='utf-8-sig')
                
                messagebox.showinfo("Sucesso", f"Dados exportados para:\n{arquivo}")
                self.status_var.set(f"CSV exportado: {arquivo}")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar CSV:\n{str(e)}")
    
    def exportar_txt(self):
        """Exporta os dados para TXT formatado"""
        if not self.candidatos:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar!")
            return
        
        arquivo = filedialog.asksaveasfilename(
            title="Salvar como TXT",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            try:
                with open(arquivo, 'w', encoding='utf-8') as f:
                    # Cabeçalho
                    f.write("=" * 120 + "\n")
                    f.write("LISTA DE CANDIDATOS - IFNMG PROCESSO 1048/2025\n".center(120))
                    f.write("=" * 120 + "\n")
                    f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write("=" * 120 + "\n\n")
                    
                    # Cabeçalho da tabela
                    cabecalho = f"{'#':<5} {'INSCRIÇÃO':<10} {'CPF':<20} {'NOME':<35} {'MODALIDADE':<15} {'NOTA':<8} {'SITUAÇÃO':<12}\n"
                    f.write(cabecalho)
                    f.write("-" * 120 + "\n")
                    
                    # Dados
                    for item in self.tree.get_children():
                        valores = self.tree.item(item)
                        idx = valores['text']
                        dados = valores['values']
                        
                        linha = f"{idx:<5} {dados[0]:<10} {dados[1]:<20} {dados[2][:35]:<35} {dados[3]:<15} {dados[4]:<8} {dados[5]:<12}\n"
                        f.write(linha)
                
                messagebox.showinfo("Sucesso", f"Dados exportados para:\n{arquivo}")
                self.status_var.set(f"TXT exportado: {arquivo}")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar TXT:\n{str(e)}")
    
    def imprimir_lista(self):
        """Prepara a lista para impressão"""
        if not self.candidatos:
            messagebox.showwarning("Aviso", "Nenhum dado para imprimir!")
            return
        
        # Criar uma nova janela para visualização de impressão
        print_window = tk.Toplevel(self.root)
        print_window.title("Visualização de Impressão")
        print_window.geometry("1000x700")
        
        # Texto para impressão
        text_widget = scrolledtext.ScrolledText(print_window, wrap=tk.WORD, font=('Courier', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Construir conteúdo
        conteudo = "=" * 100 + "\n"
        conteudo += "LISTA DE CANDIDATOS - TÉCNICO EM INTELIGÊNCIA ARTIFICIAL\n".center(100)
        conteudo += "PROCESSO SELETIVO 1048/2025 - IFNMG\n".center(100)
        conteudo += "=" * 100 + "\n\n"
        
        # Adicionar candidatos
        for item in self.tree.get_children():
            valores = self.tree.item(item)
            idx = valores['text']
            dados = valores['values']
            
            linha = f"{idx:>3}. {dados[0]} | {dados[2][:30]:30} | {dados[3]:15} | {dados[4]:>6} | {dados[5]:10}\n"
            conteudo += linha
        
        text_widget.insert(1.0, conteudo)
        
        # Botão para imprimir
        btn_frame = ttk.Frame(print_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(btn_frame, text="🖨️ Imprimir", 
                  command=lambda: self.imprimir_texto(conteudo)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Fechar", 
                  command=print_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def imprimir_texto(self, texto):
        """Função para imprimir o texto"""
        try:
            import tempfile
            import os
            
            # Criar arquivo temporário
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as tmp:
                tmp.write(texto)
                tmp_path = tmp.name
            
            # Abrir no programa padrão (que pode imprimir)
            os.startfile(tmp_path)
            
        except Exception as e:
            messagebox.showerror("Erro de Impressão", f"Não foi possível imprimir:\n{str(e)}")

def main():
    root = tk.Tk()
    app = CandidatoOrganizador(root)
    root.mainloop()

if __name__ == "__main__":
    main()