import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
import pandas as pd
from datetime import datetime
import os
import json
import shutil
from collections import defaultdict
from tkcalendar import Calendar, DateEntry

class ChromebookScheduler:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Agendamento de Chromebooks")
        self.root.geometry("950x700")
        
        # Arquivo para salvar os dados
        self.filename = "agendamentos_chromebooks.xlsx"
        self.turmas_file = "turmas_config.json"
        self.config_file = "config.json"
        
        # Quantidade total disponível
        self.total_disponivel = 100
        
        # Criar arquivo de turmas se não existir
        if not os.path.exists(self.turmas_file):
            self.create_default_turmas()
        
        # Criar arquivo Excel se não existir
        if not os.path.exists(self.filename):
            self.create_empty_file()
        
        # Carregar configurações
        self.load_config()
        
        self.setup_ui()
        self.load_data()
        self.load_turmas_list()
        self.atualizar_hud_disponiveis()
    
    def load_config(self):
        """Carrega configurações salvas"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.total_disponivel = config.get('total_disponivel', 100)
            except:
                self.total_disponivel = 100
    
    def save_config(self):
        """Salva configurações"""
        config = {
            'total_disponivel': self.total_disponivel,
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    
    def backup_data(self):
        """Cria backup do arquivo de dados"""
        if os.path.exists(self.filename):
            backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{self.filename}"
            shutil.copy2(self.filename, backup_name)
            # Manter apenas últimos 5 backups
            backups = sorted([f for f in os.listdir() if f.startswith("backup_") and f.endswith(".xlsx")])
            for old_backup in backups[:-5]:
                try:
                    os.remove(old_backup)
                except:
                    pass
    
    def create_default_turmas(self):
        """Cria arquivo de configuração com turmas padrão"""
        default_turmas = [
            "6° Ano A", "6° Ano B", "6° Ano C",
            "7° Ano A", "7° Ano B", "7° Ano C",
            "8° Ano A", "8° Ano B", "8° Ano C",
            "9° Ano A", "9° Ano B", "9° Ano C",
            "1° EM A", "1° EM B", "1° EM C",
            "2° EM A", "2° EM B", "2° EM C",
            "3° EM A", "3° EM B", "3° EM C"
        ]
        with open(self.turmas_file, 'w', encoding='utf-8') as f:
            json.dump(default_turmas, f, ensure_ascii=False, indent=2)
    
    def load_turmas_list(self):
        """Carrega a lista de turmas do arquivo JSON"""
        try:
            with open(self.turmas_file, 'r', encoding='utf-8') as f:
                turmas = json.load(f)
            self.turma_combo['values'] = turmas
            if turmas:
                self.turma_combo.set(turmas[0])
            else:
                self.turma_combo.set("")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar turmas: {str(e)}")
            self.turma_combo['values'] = []
    
    def save_turmas_list(self, turmas):
        """Salva a lista de turmas no arquivo JSON"""
        with open(self.turmas_file, 'w', encoding='utf-8') as f:
            json.dump(turmas, f, ensure_ascii=False, indent=2)
        self.load_turmas_list()
    
    def create_empty_file(self):
        """Cria um arquivo Excel vazio com as colunas necessárias"""
        df = pd.DataFrame(columns=[
            'Professor', 
            'Turma',
            'Quantidade de Chromebooks', 
            'Data de Retirada', 
            'Horário da Retirada', 
            'Horário da Devolução', 
            'Observações'
        ])
        df.to_excel(self.filename, index=False)
    
    def abrir_calendario_verificacao(self):
        """Abre um calendário popup para seleção de data (verificação)"""
        self.calendario_popup = tk.Toplevel(self.root)
        self.calendario_popup.title("Selecionar Data para Verificação")
        self.calendario_popup.geometry("500x500")
        self.calendario_popup.transient(self.root)
        self.calendario_popup.grab_set()
        
        # Centralizar na tela
        self.calendario_popup.update_idletasks()
        x = (self.calendario_popup.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.calendario_popup.winfo_screenheight() // 2) - (500 // 2)
        self.calendario_popup.geometry(f"500x500+{x}+{y}")
        
        # Frame principal
        main_frame = ttk.Frame(self.calendario_popup, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="Selecione a Data", font=('Arial', 14, 'bold'))
        titulo.pack(pady=10)
        
        # Calendário em português
        self.calendario = Calendar(main_frame, selectmode='day', date_pattern='dd/mm/yyyy',
                                    locale='pt_BR', font=('Arial', 12),
                                    background='lightblue', foreground='black',
                                    bordercolor='darkblue', headersbackground='darkblue',
                                    headersforeground='white', selectbackground='darkblue',
                                    selectforeground='white', weekendbackground='lightgray',
                                    weekendforeground='red', othermonthbackground='lightgray',
                                    othermonthforeground='gray')
        self.calendario.pack(pady=20, fill=tk.BOTH, expand=True)
        
        # Label de instrução
        instrucao = ttk.Label(main_frame, text="Clique em uma data para selecionar", 
                              font=('Arial', 10), foreground='blue')
        instrucao.pack(pady=5)
        
        # Frame de botões
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.pack(pady=20)
        
        ttk.Button(botoes_frame, text="Selecionar Data", 
                  command=self.selecionar_data_verificacao, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Cancelar", 
                  command=self.calendario_popup.destroy, width=15).pack(side=tk.LEFT, padx=10)
    
    def selecionar_data_verificacao(self):
        """Seleciona a data do calendário para verificação"""
        data_selecionada = self.calendario.get_date()
        self.verificar_data_entry.delete(0, tk.END)
        self.verificar_data_entry.insert(0, data_selecionada)
        self.calendario_popup.destroy()
        self.atualizar_hud_disponiveis()
    
    def abrir_calendario_data_retirada(self):
        """Abre um calendário popup para seleção de data de retirada"""
        self.calendario_popup2 = tk.Toplevel(self.root)
        self.calendario_popup2.title("Selecionar Data de Retirada")
        self.calendario_popup2.geometry("500x500")
        self.calendario_popup2.transient(self.root)
        self.calendario_popup2.grab_set()
        
        # Centralizar na tela
        self.calendario_popup2.update_idletasks()
        x = (self.calendario_popup2.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.calendario_popup2.winfo_screenheight() // 2) - (500 // 2)
        self.calendario_popup2.geometry(f"500x500+{x}+{y}")
        
        # Frame principal
        main_frame = ttk.Frame(self.calendario_popup2, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="Selecione a Data de Retirada", font=('Arial', 14, 'bold'))
        titulo.pack(pady=10)
        
        # Calendário em português
        self.calendario2 = Calendar(main_frame, selectmode='day', date_pattern='dd/mm/yyyy',
                                     locale='pt_BR', font=('Arial', 12),
                                     background='lightblue', foreground='black',
                                     bordercolor='darkblue', headersbackground='darkblue',
                                     headersforeground='white', selectbackground='darkblue',
                                     selectforeground='white', weekendbackground='lightgray',
                                     weekendforeground='red', othermonthbackground='lightgray',
                                     othermonthforeground='gray')
        self.calendario2.pack(pady=20, fill=tk.BOTH, expand=True)
        
        # Label de instrução
        instrucao = ttk.Label(main_frame, text="Clique em uma data para selecionar", 
                              font=('Arial', 10), foreground='blue')
        instrucao.pack(pady=5)
        
        # Frame de botões
        botoes_frame = ttk.Frame(main_frame)
        botoes_frame.pack(pady=20)
        
        ttk.Button(botoes_frame, text="Selecionar Data", 
                  command=self.selecionar_data_retirada, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(botoes_frame, text="Cancelar", 
                  command=self.calendario_popup2.destroy, width=15).pack(side=tk.LEFT, padx=10)
    
    def selecionar_data_retirada(self):
        """Seleciona a data do calendário para retirada"""
        data_selecionada = self.calendario2.get_date()
        self.data_entry.delete(0, tk.END)
        self.data_entry.insert(0, data_selecionada)
        self.calendario_popup2.destroy()
        self.atualizar_hud_disponiveis()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Agendamento de Chromebooks", 
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Frame para seleção de data e HUD
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Seleção de data para verificar disponibilidade com botão de calendário
        ttk.Label(top_frame, text="Verificar disponibilidade para data:").pack(side=tk.LEFT, padx=5)
        
        frame_data_verificar = ttk.Frame(top_frame)
        frame_data_verificar.pack(side=tk.LEFT, padx=5)
        
        self.verificar_data_entry = ttk.Entry(frame_data_verificar, width=15, font=('Arial', 10))
        self.verificar_data_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.verificar_data_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        btn_calendario_verificar = ttk.Button(frame_data_verificar, text="📅", width=3, 
                                               command=self.abrir_calendario_verificacao)
        btn_calendario_verificar.pack(side=tk.LEFT)
        
        ttk.Button(top_frame, text="Verificar", command=self.atualizar_hud_disponiveis).pack(side=tk.LEFT, padx=5)
        
        # HUD de disponibilidade
        hud_frame = ttk.LabelFrame(main_frame, text="Status de Disponibilidade", padding="10")
        hud_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.disponivel_label = ttk.Label(hud_frame, text=f"Chromebooks Disponíveis: {self.total_disponivel}", 
                                          font=('Arial', 12, 'bold'), foreground='green')
        self.disponivel_label.pack()
        
        # Formulário de cadastro
        form_frame = ttk.LabelFrame(main_frame, text="Novo Agendamento", padding="10")
        form_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Campos do formulário
        # Professor
        ttk.Label(form_frame, text="Professor:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.professor_entry = ttk.Entry(form_frame, width=30)
        self.professor_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 20))
        
        # Turma com botão de edição
        ttk.Label(form_frame, text="Turma:").grid(row=0, column=2, sticky=tk.W, pady=5)
        turma_frame = ttk.Frame(form_frame)
        turma_frame.grid(row=0, column=3, sticky=tk.W, pady=5, padx=(0, 20))
        
        self.turma_combo = ttk.Combobox(turma_frame, width=20)
        self.turma_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(turma_frame, text="✎", width=3, command=self.manage_turmas).pack(side=tk.LEFT)
        
        # Quantidade
        ttk.Label(form_frame, text="Quantidade de Chromebooks:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.quantidade_spinbox = ttk.Spinbox(form_frame, from_=1, to=self.total_disponivel, width=10)
        self.quantidade_spinbox.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Data de retirada com botão de calendário
        ttk.Label(form_frame, text="Data de Retirada:").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        frame_data_retirada = ttk.Frame(form_frame)
        frame_data_retirada.grid(row=1, column=3, sticky=(tk.W, tk.E), pady=5)
        
        self.data_entry = ttk.Entry(frame_data_retirada, width=25, font=('Arial', 10))
        self.data_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.data_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        btn_calendario_retirada = ttk.Button(frame_data_retirada, text="📅", width=3, 
                                              command=self.abrir_calendario_data_retirada)
        btn_calendario_retirada.pack(side=tk.LEFT)
        
        # Horário de retirada
        ttk.Label(form_frame, text="Horário da Retirada:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.horario_retirada_combo = ttk.Combobox(form_frame, values=self.generate_time_options(), width=15)
        self.horario_retirada_combo.grid(row=2, column=1, sticky=tk.W, pady=5)
        self.horario_retirada_combo.set("07:00")
        
        # Horário de devolução
        ttk.Label(form_frame, text="Horário da Devolução:").grid(row=2, column=2, sticky=tk.W, pady=5)
        self.horario_devolucao_combo = ttk.Combobox(form_frame, values=self.generate_time_options(), width=15)
        self.horario_devolucao_combo.grid(row=2, column=3, sticky=tk.W, pady=5)
        self.horario_devolucao_combo.set("12:20")
        
        # Observações
        ttk.Label(form_frame, text="Observações:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.observacoes_text = scrolledtext.ScrolledText(form_frame, width=30, height=4)
        self.observacoes_text.grid(row=3, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=(0, 20))
        
        # Botões
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=4, column=0, columnspan=4, pady=10)
        
        ttk.Button(button_frame, text="Adicionar Agendamento", command=self.add_schedule).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpar Formulário", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remover Selecionado", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Atualizar Tabela", command=self.load_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exportar Relatório", command=self.exportar_relatorio).pack(side=tk.LEFT, padx=5)
        
        # Filtros de pesquisa
        filter_frame = ttk.LabelFrame(main_frame, text="Filtros", padding="5")
        filter_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(filter_frame, text="Filtrar por Professor:").pack(side=tk.LEFT, padx=5)
        self.filter_entry = ttk.Entry(filter_frame, width=20)
        self.filter_entry.pack(side=tk.LEFT, padx=5)
        self.filter_entry.bind('<KeyRelease>', lambda e: self.apply_filter())
        
        ttk.Label(filter_frame, text="Data:").pack(side=tk.LEFT, padx=5)
        self.filter_date = ttk.Entry(filter_frame, width=12)
        self.filter_date.pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="Aplicar Filtro", 
                   command=self.apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="Limpar Filtro", 
                   command=self.clear_filter).pack(side=tk.LEFT, padx=5)
        
        # Tabela de agendamentos
        table_frame = ttk.LabelFrame(main_frame, text="Agendamentos Realizados", padding="10")
        table_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Scrollbars
        scrollbar_y = ttk.Scrollbar(table_frame)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview para mostrar os dados
        self.tree = ttk.Treeview(table_frame, 
                                 columns=('Professor', 'Turma', 'Quantidade', 'Data', 'Hora Retirada', 'Hora Devolução', 'Observações'),
                                 show='headings',
                                 yscrollcommand=scrollbar_y.set,
                                 xscrollcommand=scrollbar_x.set)
        
        # Configurar colunas
        self.tree.heading('Professor', text='Professor')
        self.tree.heading('Turma', text='Turma')
        self.tree.heading('Quantidade', text='Quantidade')
        self.tree.heading('Data', text='Data Retirada')
        self.tree.heading('Hora Retirada', text='Hora Retirada')
        self.tree.heading('Hora Devolução', text='Hora Devolução')
        self.tree.heading('Observações', text='Observações')
        
        self.tree.column('Professor', width=150)
        self.tree.column('Turma', width=100)
        self.tree.column('Quantidade', width=80)
        self.tree.column('Data', width=100)
        self.tree.column('Hora Retirada', width=100)
        self.tree.column('Hora Devolução', width=100)
        self.tree.column('Observações', width=200)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # Bind de seleção
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Barra de status
        self.status_label = ttk.Label(main_frame, text="Pronto", relief=tk.SUNKEN)
        self.status_label.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Variável para armazenar dados originais (para filtro)
        self.df_original = None
    
    def apply_filter(self):
        """Aplica filtros na tabela"""
        professor_filter = self.filter_entry.get().strip().lower()
        data_filter = self.filter_date.get().strip()
        
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            df = pd.read_excel(self.filename)
            if df.empty:
                return
            
            # Aplicar filtros
            if professor_filter:
                df = df[df['Professor'].str.lower().str.contains(professor_filter, na=False)]
            
            if data_filter:
                df = df[df['Data de Retirada'] == data_filter]
            
            # Ordenar por data
            if not df.empty:
                df['Data para ordenar'] = pd.to_datetime(df['Data de Retirada'], format='%d/%m/%Y', errors='coerce')
                df = df.sort_values('Data para ordenar').drop('Data para ordenar', axis=1)
            
            # Inserir dados na tabela
            for idx, row in df.iterrows():
                self.tree.insert('', 'end', values=(
                    row['Professor'],
                    row['Turma'],
                    row['Quantidade de Chromebooks'],
                    row['Data de Retirada'],
                    row['Horário da Retirada'],
                    row['Horário da Devolução'],
                    row['Observações']
                ))
            
            self.status_label.config(text=f"Total de agendamentos: {len(df)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao aplicar filtro: {str(e)}")
    
    def clear_filter(self):
        """Limpa os filtros aplicados"""
        self.filter_entry.delete(0, tk.END)
        self.filter_date.delete(0, tk.END)
        self.load_data()
    
    def exportar_relatorio(self):
        """Exporta relatório em CSV ou Excel"""
        from tkinter import filedialog
        
        filetype = [('CSV files', '*.csv'), ('Excel files', '*.xlsx')]
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=filetype
        )
        
        if filename:
            try:
                df = pd.read_excel(self.filename)
                if filename.endswith('.csv'):
                    df.to_csv(filename, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(filename, index=False)
                messagebox.showinfo("Sucesso", f"Relatório exportado para {filename}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
    
    def atualizar_hud_disponiveis(self):
        """Atualiza o HUD com a quantidade disponível para a data selecionada"""
        # Obter a data para verificar disponibilidade
        data_verificar = self.verificar_data_entry.get().strip()
        
        if not data_verificar:
            self.disponivel_label.config(text=f"Chromebooks Disponíveis: {self.total_disponivel}")
            self.quantidade_spinbox.config(to=self.total_disponivel)
            return self.total_disponivel
        
        try:
            # Validar formato da data
            datetime.strptime(data_verificar, "%d/%m/%Y")
            
            df = pd.read_excel(self.filename)
            
            if df.empty:
                disponivel = self.total_disponivel
            else:
                # Filtrar apenas agendamentos da data selecionada
                df_filtrado = df[df['Data de Retirada'] == data_verificar]
                total_agendado = df_filtrado['Quantidade de Chromebooks'].sum() if not df_filtrado.empty else 0
                disponivel = self.total_disponivel - total_agendado
            
            self.disponivel_label.config(text=f"Chromebooks Disponíveis para {data_verificar}: {disponivel}")
            
            # Atualizar o limite máximo do spinbox
            self.quantidade_spinbox.config(to=disponivel if disponivel > 0 else 1)
            
            if disponivel <= 10:
                self.disponivel_label.config(foreground='red')
            elif disponivel <= 30:
                self.disponivel_label.config(foreground='orange')
            else:
                self.disponivel_label.config(foreground='green')
            
            return disponivel
        except ValueError:
            self.disponivel_label.config(text=f"Data inválida! Use DD/MM/AAAA")
            return self.total_disponivel
        except Exception as e:
            self.disponivel_label.config(text=f"Chromebooks Disponíveis: {self.total_disponivel}")
            # Registrar log do erro
            with open("error_log.txt", "a", encoding='utf-8') as log_file:
                log_file.write(f"{datetime.now()}: {str(e)}\n")
            return self.total_disponivel
    
    def manage_turmas(self):
        """Abre janela para gerenciar a lista de turmas"""
        # Verificar se a janela já existe
        if hasattr(self, 'manage_window') and self.manage_window.winfo_exists():
            self.manage_window.lift()
            return
        
        self.manage_window = tk.Toplevel(self.root)
        self.manage_window.title("Gerenciar Turmas")
        self.manage_window.geometry("500x500")
        self.manage_window.transient(self.root)
        self.manage_window.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(self.manage_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Lista de turmas
        ttk.Label(main_frame, text="Lista de Turmas:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Frame para lista e scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.turmas_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Arial', 10))
        self.turmas_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.turmas_listbox.yview)
        
        # Carregar turmas na listbox
        try:
            with open(self.turmas_file, 'r', encoding='utf-8') as f:
                turmas = json.load(f)
                for turma in turmas:
                    self.turmas_listbox.insert(tk.END, turma)
        except:
            pass
        
        # Frame para botões de gerenciamento
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="Adicionar Turma", command=self.add_turma).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Editar Turma", command=self.edit_turma).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remover Turma", command=self.remove_turma).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Subir", command=self.move_up_turma).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Descer", command=self.move_down_turma).pack(side=tk.LEFT, padx=5)
        
        # Botão fechar
        ttk.Button(main_frame, text="Fechar", command=self.manage_window.destroy).pack()
    
    def add_turma(self):
        """Adiciona uma nova turma"""
        nova_turma = simpledialog.askstring("Adicionar Turma", "Digite o nome da nova turma:")
        if nova_turma and nova_turma.strip():
            nova_turma = nova_turma.strip()
            # Verificar se já existe
            turmas_atual = list(self.turmas_listbox.get(0, tk.END))
            if nova_turma not in turmas_atual:
                self.turmas_listbox.insert(tk.END, nova_turma)
                self.save_turmas_from_listbox()
                messagebox.showinfo("Sucesso", f"Turma '{nova_turma}' adicionada com sucesso!")
            else:
                messagebox.showwarning("Aviso", "Esta turma já existe na lista!")
    
    def edit_turma(self):
        """Edita a turma selecionada"""
        selecionado = self.turmas_listbox.curselection()
        if selecionado:
            turma_atual = self.turmas_listbox.get(selecionado[0])
            nova_turma = simpledialog.askstring("Editar Turma", "Editar nome da turma:", initialvalue=turma_atual)
            if nova_turma and nova_turma.strip():
                nova_turma = nova_turma.strip()
                # Verificar se já existe (exceto a própria)
                turmas_atual = list(self.turmas_listbox.get(0, tk.END))
                if nova_turma not in turmas_atual or nova_turma == turma_atual:
                    self.turmas_listbox.delete(selecionado[0])
                    self.turmas_listbox.insert(selecionado[0], nova_turma)
                    self.save_turmas_from_listbox()
                    messagebox.showinfo("Sucesso", f"Turma alterada de '{turma_atual}' para '{nova_turma}'!")
                else:
                    messagebox.showwarning("Aviso", "Esta turma já existe na lista!")
        else:
            messagebox.showwarning("Aviso", "Selecione uma turma para editar!")
    
    def remove_turma(self):
        """Remove a turma selecionada"""
        selecionado = self.turmas_listbox.curselection()
        if selecionado:
            turma = self.turmas_listbox.get(selecionado[0])
            if messagebox.askyesno("Confirmar", f"Tem certeza que deseja remover a turma '{turma}'?"):
                self.turmas_listbox.delete(selecionado[0])
                self.save_turmas_from_listbox()
                messagebox.showinfo("Sucesso", f"Turma '{turma}' removida com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Selecione uma turma para remover!")
    
    def move_up_turma(self):
        """Move a turma selecionada para cima"""
        selecionado = self.turmas_listbox.curselection()
        if selecionado and selecionado[0] > 0:
            index = selecionado[0]
            turma = self.turmas_listbox.get(index)
            self.turmas_listbox.delete(index)
            self.turmas_listbox.insert(index - 1, turma)
            self.turmas_listbox.selection_set(index - 1)
            self.save_turmas_from_listbox()
    
    def move_down_turma(self):
        """Move a turma selecionada para baixo"""
        selecionado = self.turmas_listbox.curselection()
        if selecionado and selecionado[0] < self.turmas_listbox.size() - 1:
            index = selecionado[0]
            turma = self.turmas_listbox.get(index)
            self.turmas_listbox.delete(index)
            self.turmas_listbox.insert(index + 1, turma)
            self.turmas_listbox.selection_set(index + 1)
            self.save_turmas_from_listbox()
    
    def save_turmas_from_listbox(self):
        """Salva as turmas da listbox no arquivo JSON"""
        turmas = list(self.turmas_listbox.get(0, tk.END))
        self.save_turmas_list(turmas)
    
    def generate_time_options(self):
        """Gera lista de horários pré-definida"""
        return ["07:00", "07:50", "08:40", "09:00", "09:50", "10:40", "11:30", "12:20", "13:10", "14:00"]
    
    def add_schedule(self):
        """Adiciona um novo agendamento"""
        # Validar campos
        professor = self.professor_entry.get().strip()
        turma = self.turma_combo.get()
        quantidade = self.quantidade_spinbox.get()
        data = self.data_entry.get().strip()
        hora_retirada = self.horario_retirada_combo.get()
        hora_devolucao = self.horario_devolucao_combo.get()
        observacoes = self.observacoes_text.get("1.0", tk.END).strip()
        
        if not professor:
            messagebox.showerror("Erro", "O campo Professor é obrigatório!")
            return
        
        if not turma:
            messagebox.showerror("Erro", "Selecione uma turma!")
            return
        
        if not quantidade or not quantidade.isdigit():
            messagebox.showerror("Erro", "Quantidade inválida!")
            return
        
        quantidade_int = int(quantidade)
        
        # Verificar disponibilidade para a data específica
        try:
            datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Data inválida! Use o formato DD/MM/AAAA")
            return
        
        # Calcular disponibilidade para a data
        try:
            df = pd.read_excel(self.filename)
            if not df.empty:
                df_filtrado = df[df['Data de Retirada'] == data]
                total_agendado = df_filtrado['Quantidade de Chromebooks'].sum() if not df_filtrado.empty else 0
                disponivel = self.total_disponivel - total_agendado
            else:
                disponivel = self.total_disponivel
        except:
            disponivel = self.total_disponivel
        
        if quantidade_int > disponivel:
            messagebox.showerror("Erro", f"Quantidade indisponível para a data {data}! Apenas {disponivel} chromebooks disponíveis.")
            return
        
        # Criar DataFrame com o novo registro
        new_record = pd.DataFrame([{
            'Professor': professor,
            'Turma': turma,
            'Quantidade de Chromebooks': quantidade_int,
            'Data de Retirada': data,
            'Horário da Retirada': hora_retirada,
            'Horário da Devolução': hora_devolucao,
            'Observações': observacoes
        }])
        
        try:
            # Criar backup antes de adicionar
            self.backup_data()
            
            # Carregar dados existentes
            df = pd.read_excel(self.filename)
            # Adicionar novo registro
            df = pd.concat([df, new_record], ignore_index=True)
            # Salvar
            df.to_excel(self.filename, index=False)
            
            self.status_label.config(text=f"Agendamento adicionado com sucesso para {professor} - Turma {turma}")
            self.load_data()
            self.atualizar_hud_disponiveis()
            self.clear_form()
            messagebox.showinfo("Sucesso", "Agendamento realizado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {str(e)}")
            # Registrar log do erro
            with open("error_log.txt", "a", encoding='utf-8') as log_file:
                log_file.write(f"{datetime.now()}: {str(e)}\n")
    
    def load_data(self):
        """Carrega os dados da planilha para a tabela com tratamento de erro melhorado"""
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            df = pd.read_excel(self.filename)
            
            # Verificar se as colunas existem
            expected_columns = ['Professor', 'Turma', 'Quantidade de Chromebooks', 
                               'Data de Retirada', 'Horário da Retirada', 
                               'Horário da Devolução', 'Observações']
            
            for col in expected_columns:
                if col not in df.columns:
                    df[col] = ""  # Adicionar coluna faltante
            
            # Ordenar por data
            if not df.empty:
                df['Data para ordenar'] = pd.to_datetime(df['Data de Retirada'], format='%d/%m/%Y', errors='coerce')
                df = df.sort_values('Data para ordenar').drop('Data para ordenar', axis=1)
            
            # Inserir dados na tabela
            for idx, row in df.iterrows():
                self.tree.insert('', 'end', values=(
                    row['Professor'],
                    row['Turma'],
                    row['Quantidade de Chromebooks'],
                    row['Data de Retirada'],
                    row['Horário da Retirada'],
                    row['Horário da Devolução'],
                    row['Observações']
                ))
            
            self.status_label.config(text=f"Total de agendamentos: {len(df)}")
        except FileNotFoundError:
            self.create_empty_file()
            self.load_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            # Registrar log do erro
            with open("error_log.txt", "a", encoding='utf-8') as log_file:
                log_file.write(f"{datetime.now()}: {str(e)}\n")
    
    def clear_form(self):
        """Limpa o formulário"""
        self.professor_entry.delete(0, tk.END)
        self.load_turmas_list()  # Recarrega a lista de turmas
        self.quantidade_spinbox.set("1")
        self.data_entry.delete(0, tk.END)
        self.data_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.horario_retirada_combo.set("07:00")
        self.horario_devolucao_combo.set("12:20")
        self.observacoes_text.delete("1.0", tk.END)
        self.professor_entry.focus()
        self.atualizar_hud_disponiveis()
    
    def delete_selected(self):
        """Remove o agendamento selecionado"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um agendamento para remover!")
            return
        
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja remover o agendamento selecionado?"):
            # Obter valores do item selecionado (apenas o primeiro)
            values = self.tree.item(selected[0])['values']
            
            try:
                # Criar backup antes de remover
                self.backup_data()
                
                df = pd.read_excel(self.filename)
                # Encontrar e remover o registro
                mask = (df['Professor'] == values[0]) & \
                       (df['Turma'] == values[1]) & \
                       (df['Quantidade de Chromebooks'] == int(values[2])) & \
                       (df['Data de Retirada'] == values[3]) & \
                       (df['Horário da Retirada'] == values[4])
                
                df = df[~mask]
                df.to_excel(self.filename, index=False)
                
                self.load_data()
                self.atualizar_hud_disponiveis()
                self.status_label.config(text="Agendamento removido com sucesso!")
                messagebox.showinfo("Sucesso", "Agendamento removido com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao remover: {str(e)}")
                # Registrar log do erro
                with open("error_log.txt", "a", encoding='utf-8') as log_file:
                    log_file.write(f"{datetime.now()}: {str(e)}\n")
    
    def on_select(self, event):
        """Preenche o formulário com o agendamento selecionado"""
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0])['values']
            self.professor_entry.delete(0, tk.END)
            self.professor_entry.insert(0, values[0])
            self.turma_combo.set(values[1])
            self.quantidade_spinbox.set(values[2])
            self.data_entry.delete(0, tk.END)
            self.data_entry.insert(0, values[3])
            self.horario_retirada_combo.set(values[4])
            self.horario_devolucao_combo.set(values[5])
            self.observacoes_text.delete("1.0", tk.END)
            self.observacoes_text.insert("1.0", values[6])
            self.atualizar_hud_disponiveis()


if __name__ == "__main__":
    # Instalar dependências necessárias
    try:
        import pandas
        import openpyxl
        from tkcalendar import Calendar, DateEntry
    except ImportError:
        print("Instalando dependências necessárias...")
        import subprocess
        import sys
        
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl", "tkcalendar"])
        print("Dependências instaladas com sucesso!")
    
    root = tk.Tk()
    app = ChromebookScheduler(root)
    root.mainloop()