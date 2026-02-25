import os
import sys
import time

import pandas as pd
import pythoncom
import win32com.client
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (QApplication, QFileDialog, QHBoxLayout, QLabel,
                               QMainWindow, QMessageBox, QProgressBar,
                               QPushButton, QSplitter, QTextEdit, QVBoxLayout,
                               QWidget)


class AutocadWorker(QThread):
    progress = Signal(int)
    finished = Signal(dict)
    error = Signal(str)
    log = Signal(str)
    
    def __init__(self, excel_path, template_folder):
        super().__init__()
        self.excel_path = excel_path
        self.template_folder = template_folder
        
    def run(self):
        try:
            # Estatísticas para o relatório final
            stats = {
                "total": 0,
                "success": 0,
                "template_not_found": 0,
                "errors": 0,
                "no_attributes": 0,  # Nova categoria para arquivos sem atributos
                "error_details": [],
                "not_found_details": [],
                "no_attributes_details": []  # Lista para arquivos sem atributos
            }
            
            # Inicializa COM na thread atual
            pythoncom.CoInitialize()
            
            # Lê o arquivo Excel
            self.log.emit("Lendo arquivo Excel...")
            df = pd.read_excel(self.excel_path)
            
            # Verifica se todas as colunas necessárias existem
            required_columns = ['Posicao', 'TipoSuporte', 'Elevacao', 'H', 'L', 'M']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                self.error.emit(f"Colunas faltando no Excel: {', '.join(missing_columns)}")
                stats["errors"] += 1
                stats["error_details"].append(f"Colunas faltando: {', '.join(missing_columns)}")
                self.finished.emit(stats)
                return
            
            # Inicializa o AutoCAD
            self.log.emit("Conectando ao AutoCAD...")
            try:
                acad = win32com.client.Dispatch("AutoCAD.Application")
                acad.Visible = True
                self.log.emit("Conexão com AutoCAD estabelecida com sucesso.")
            except Exception as e:
                self.error.emit(f"Erro ao conectar com AutoCAD: {str(e)}")
                stats["errors"] += 1
                stats["error_details"].append(f"Falha na conexão com AutoCAD: {str(e)}")
                self.finished.emit(stats)
                return
                
            # Processa cada linha
            total_rows = len(df)
            stats["total"] = total_rows
            self.log.emit(f"Iniciando processamento de {total_rows} registros.")
            
            for i, row in df.iterrows():
                posicao = str(row['Posicao'])
                tipo_suporte = str(row['TipoSuporte'])
                elevacao = str(row['Elevacao']).replace(',', '.')
                h = str(row['H']) if pd.notna(row['H']) else "-"
                l = str(row['L']) if pd.notna(row['L']) else "-"
                m = str(row['M']) if pd.notna(row['M']) else "-"
                
                # Atualiza a barra de progresso
                progress_percent = int((i + 1) / total_rows * 100)
                self.progress.emit(progress_percent)
                
                # Log detalhado para o usuário
                self.log.emit(f"[{i+1}/{total_rows}] Processando {posicao} (Tipo: {tipo_suporte})")
                
                # Construir caminho do template
                template_path = os.path.join(self.template_folder, f"{tipo_suporte}.dwg")
                output_path = os.path.join(os.path.dirname(self.excel_path), f"{posicao}.dwg")
                
                # Verifica se o template existe
                if not os.path.exists(template_path):
                    self.log.emit(f"  ⚠️ Template para {tipo_suporte} não encontrado. Pulando...")
                    stats["template_not_found"] += 1
                    stats["not_found_details"].append(f"{posicao} (Tipo: {tipo_suporte})")
                    continue
                
                retry_count = 3  # Número de tentativas
                for attempt in range(retry_count):
                    try:
                        # Abre o template
                        self.log.emit(f"  - Abrindo template: {tipo_suporte}.dwg (Tentativa {attempt + 1})")
                        doc = acad.Documents.Open(template_path)
                        time.sleep(1.0)  # Aumentado para 1 segundo para garantir que o arquivo foi aberto corretamente
                        
                        # Modifica os atributos de texto
                        self.log.emit(f"  - Verificando atributos...")
                        attr_count = 0
                        found_attributes = False
                        
                        # Verifica primeiro se há blocos com atributos
                        for entity in doc.PaperSpace:
                            if entity.ObjectName == "AcDbBlockReference" and entity.HasAttributes:
                                found_attributes = True
                                break
                        
                        if not found_attributes:
                            self.log.emit(f"  ⚠️ Nenhum bloco com atributos encontrado no template {tipo_suporte}.dwg")
                            stats["no_attributes"] += 1
                            stats["no_attributes_details"].append(f"{posicao} (Tipo: {tipo_suporte})")
                            # Fecha o documento sem salvar
                            doc.Close(False)
                            break
                        
                        # Preenche os atributos, agora que sabemos que existem
                        self.log.emit(f"  - Preenchendo atributos...")
                        for entity in doc.PaperSpace:
                            if entity.ObjectName == "AcDbBlockReference" and entity.HasAttributes:
                                for attrib in entity.GetAttributes():
                                    tag = attrib.TagString.upper()
                                    
                                    if tag == "POSICAO":
                                        attrib.TextString = posicao
                                        attr_count += 1
                                    elif tag == "TIPOSUPORTE":
                                        attrib.TextString = tipo_suporte
                                        attr_count += 1
                                    elif tag == "ELEVACAO":
                                        attrib.TextString = elevacao
                                        attr_count += 1
                                    elif tag == "H":
                                        attrib.TextString = h
                                        attr_count += 1
                                    elif tag == "L":
                                        attrib.TextString = l
                                        attr_count += 1
                                    elif tag == "M":
                                        attrib.TextString = m
                                        attr_count += 1
                        
                        # Verifica se algum atributo foi preenchido
                        if attr_count == 0:
                            self.log.emit(f"  ⚠️ Nenhum atributo correspondente encontrado no template {tipo_suporte}.dwg")
                            stats["no_attributes"] += 1
                            stats["no_attributes_details"].append(f"{posicao} (Tipo: {tipo_suporte})")
                            # Fecha o documento sem salvar
                            doc.Close(False)
                            break
                        
                        self.log.emit(f"  - {attr_count} atributos preenchidos")
                        
                        # Salva como novo arquivo apenas se atributos foram preenchidos
                        self.log.emit(f"  - Salvando como: {posicao}.dwg")
                        doc.SaveAs(output_path)
                        doc.Close()
                        self.log.emit(f"  ✅ {posicao}.dwg criado com sucesso!")
                        stats["success"] += 1
                        break  # Sai do loop de retry se bem-sucedido
                        
                    except Exception as e:
                        self.log.emit(f"  ❌ Erro ao processar {posicao} (Tentativa {attempt + 1}): {str(e)}")
                        if attempt == retry_count - 1:  # Se for a última tentativa
                            stats["errors"] += 1
                            stats["error_details"].append(f"{posicao}: {str(e)}")
                            # Tenta fechar o documento se houver erro
                            try:
                                if 'doc' in locals():
                                    doc.Close(False)
                            except:
                                pass
            
            self.log.emit("\n===== PROCESSAMENTO CONCLUÍDO =====")
            self.finished.emit(stats)
            
        except Exception as e:
            self.error.emit(f"Erro geral: {str(e)}")
            stats = {
                "total": 0,
                "success": 0,
                "template_not_found": 0,
                "no_attributes": 0,
                "errors": 1,
                "error_details": [f"Erro geral: {str(e)}"],
                "not_found_details": [],
                "no_attributes_details": []
            }
            self.finished.emit(stats)
        finally:
            # Limpa COM
            pythoncom.CoUninitialize()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Integração Excel-AutoCAD")
        self.setGeometry(100, 100, 800, 600)
        
        # Layout principal
        main_layout = QVBoxLayout()
        
        # Área superior com controles
        controls_layout = QVBoxLayout()
        
        # Área de seleção de arquivo Excel
        excel_layout = QHBoxLayout()
        self.excel_label = QLabel("Nenhum arquivo Excel selecionado")
        self.excel_button = QPushButton("Selecionar Arquivo Excel")
        self.excel_button.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.excel_label, 1)
        excel_layout.addWidget(self.excel_button, 0)
        
        # Área de seleção da pasta de templates
        template_layout = QHBoxLayout()
        self.template_label = QLabel("Nenhuma pasta de templates selecionada")
        self.template_button = QPushButton("Selecionar Pasta de Templates")
        self.template_button.clicked.connect(self.select_template_folder)
        template_layout.addWidget(self.template_label, 1)
        template_layout.addWidget(self.template_button, 0)
        
        # Botão de processamento
        self.process_button = QPushButton("Processar Dados")
        self.process_button.clicked.connect(self.process_data)
        self.process_button.setEnabled(False)
        
        # Barra de progresso
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(QLabel("Progresso:"))
        progress_layout.addWidget(self.progress_bar, 1)
        
        # Adiciona widgets ao layout de controles
        controls_layout.addLayout(excel_layout)
        controls_layout.addLayout(template_layout)
        controls_layout.addWidget(self.process_button)
        controls_layout.addLayout(progress_layout)
        
        controls_widget = QWidget()
        controls_widget.setLayout(controls_layout)
        
        # Área de log
        log_layout = QVBoxLayout()
        log_label = QLabel("Log de Processamento:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        # Usar fonte monoespaçada para melhor legibilidade do log
        font = QFont("Consolas" if sys.platform == "win32" else "Monospace")
        font.setPointSize(10)
        self.log_text.setFont(font)
        
        log_layout.addWidget(log_label)
        log_layout.addWidget(self.log_text)
        
        log_widget = QWidget()
        log_widget.setLayout(log_layout)
        
        # Criando splitter para redimensionamento
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(controls_widget)
        splitter.addWidget(log_widget)
        splitter.setSizes([200, 400])  # Tamanho inicial das seções
        
        # Adiciona o splitter ao layout principal
        main_layout.addWidget(splitter)
        
        # Widget central
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # Variáveis de instância
        self.excel_path = None
        self.template_folder = None
        self.worker = None
    
    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx *.xls)"
        )
        
        if file_path:
            self.excel_path = file_path
            self.excel_label.setText(f"Excel: {os.path.basename(file_path)}")
            self.check_ready()
    
    def select_template_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, "Selecionar Pasta de Templates"
        )
        
        if folder_path:
            self.template_folder = folder_path
            self.template_label.setText(f"Templates: {os.path.basename(folder_path)}")
            self.check_ready()
    
    def check_ready(self):
        if self.excel_path and self.template_folder:
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)
    
    def process_data(self):
        # Limpa o log
        self.log_text.clear()
        
        # Desativa os botões durante o processamento
        self.excel_button.setEnabled(False)
        self.template_button.setEnabled(False)
        self.process_button.setEnabled(False)
        
        # Adiciona cabeçalho ao log
        self.add_to_log(f"=== INÍCIO DO PROCESSAMENTO ===")
        self.add_to_log(f"Data/Hora: {time.strftime('%d/%m/%Y %H:%M:%S')}")
        self.add_to_log(f"Arquivo Excel: {self.excel_path}")
        self.add_to_log(f"Pasta de Templates: {self.template_folder}")
        self.add_to_log("-" * 50)
        
        # Cria e inicia o worker thread
        self.worker = AutocadWorker(self.excel_path, self.template_folder)
        self.worker.progress.connect(self.update_progress)
        self.worker.log.connect(self.add_to_log)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(self.processing_finished)
        self.worker.start()
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def add_to_log(self, message):
        self.log_text.append(message)
        # Rola para o fim do texto
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def show_error(self, message):
        QMessageBox.critical(self, "Erro", message)
        self.add_to_log(f"❌ ERRO: {message}")
    
    def processing_finished(self, stats):
        # Gera relatório final
        self.add_to_log("\n" + "=" * 50)
        self.add_to_log("RELATÓRIO FINAL DE PROCESSAMENTO")
        self.add_to_log("-" * 50)
        self.add_to_log(f"Total de registros processados: {stats['total']}")
        self.add_to_log(f"✅ Arquivos criados com sucesso: {stats['success']}")
        self.add_to_log(f"⚠️ Templates não encontrados: {stats['template_not_found']}")
        self.add_to_log(f"⚠️ Templates sem atributos: {stats.get('no_attributes', 0)}")
        self.add_to_log(f"❌ Erros durante o processamento: {stats['errors']}")
        
        # Detalhes sobre templates não encontrados
        if stats["template_not_found"] > 0:
            self.add_to_log("\nDetalhes de Templates Não Encontrados:")
            for detail in stats["not_found_details"]:
                self.add_to_log(f"  - {detail}")
        
        # Detalhes sobre templates sem atributos
        if stats.get("no_attributes", 0) > 0:
            self.add_to_log("\nDetalhes de Templates Sem Atributos:")
            for detail in stats["no_attributes_details"]:
                self.add_to_log(f"  - {detail}")
        
        # Detalhes sobre erros
        if stats["errors"] > 0:
            self.add_to_log("\nDetalhes de Erros:")
            for detail in stats["error_details"]:
                self.add_to_log(f"  - {detail}")
        
        self.add_to_log("\n" + "=" * 50)
        self.add_to_log(f"Processamento finalizado em: {time.strftime('%d/%m/%Y %H:%M:%S')}")
        
        # Reativa os botões
        self.excel_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.process_button.setEnabled(True)
        
        # Mostra resumo em uma mensagem
        QMessageBox.information(
            self, 
            "Processamento Concluído", 
            f"Processamento concluído!\n\n"
            f"Total: {stats['total']}\n"
            f"Sucesso: {stats['success']}\n"
            f"Templates não encontrados: {stats['template_not_found']}\n"
            f"Templates sem atributos: {stats.get('no_attributes', 0)}\n"
            f"Erros: {stats['errors']}\n\n"
            f"Veja o log para mais detalhes."
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
