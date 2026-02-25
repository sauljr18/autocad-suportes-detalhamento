import os
import sys
import time
from datetime import datetime

import pandas as pd
import pythoncom
import win32com.client
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (QApplication, QCheckBox, QFileDialog, QHBoxLayout,
                               QLabel, QMainWindow, QMessageBox, QProgressBar,
                               QPushButton, QSplitter, QTextEdit, QVBoxLayout,
                               QWidget)


class ProcessingConfig:
    """Configurações centralizadas do processamento."""
    # Tentativas de retry
    RETRY_COUNT = 3

    # Tempos de sleep (em segundos) - OTIMIZADOS para performance
    SLEEP_AFTER_OPEN = 1.0       # Reduzido de 2.0 para 1.0
    SLEEP_AFTER_SAVE = 0.3       # Reduzido de 0.5 para 0.3
    SLEEP_BETWEEN_OPS = 0.15     # Reduzido de 0.3 para 0.15

    # Colunas obrigatórias do Excel
    REQUIRED_COLUMNS = ['POSICAO', 'TipoSuporte', 'Elevacao', 'MEDIDA_H', 'MEDIDA_L', 'MEDIDA_M', 'MEDIDA_H1', 'MEDIDA_H2', 'MEDIDA_L1', 'MEDIDA_L2', 'MEDIDA_B']

    # Mapeamento de atributos do AutoCAD
    ATTRIBUTE_TAGS = ["POSICAO", "TIPOSUPORTE", "ELEVACAO", "H", "L", "M", "H1", "H2", "L1", "L2", "B", "DATA_ATUAL"]


class ProcessingStats:
    """Gerencia estatísticas do processamento."""
    def __init__(self):
        self.total = 0
        self.success = 0
        self.template_not_found = 0
        self.errors = 0
        self.no_attributes = 0
        self.duplicates = 0
        self.error_details = []
        self.not_found_details = []
        self.no_attributes_details = []
        self.duplicate_details = []

    def to_dict(self):
        """Converte para dicionário para sinal."""
        return {
            "total": self.total,
            "success": self.success,
            "template_not_found": self.template_not_found,
            "errors": self.errors,
            "no_attributes": self.no_attributes,
            "duplicates": self.duplicates,
            "error_details": self.error_details,
            "not_found_details": self.not_found_details,
            "no_attributes_details": self.no_attributes_details,
            "duplicate_details": self.duplicate_details
        }


class COMErrorHandler:
    """Handler para erros COM RPC_E_CALL_REJECTED."""

    # HRESULT para RPC_E_CALL_REJECTED
    RPC_E_CALL_REJECTED = -2147418111

    @staticmethod
    def execute_with_retry(func, operation_name="Operation", max_retries=3):
        """
        Executa uma função com retry para erros RPC_E_CALL_REJECTED.

        Args:
            func: Função a executar
            operation_name: Nome da operação para log
            max_retries: Número máximo de tentativas

        Returns:
            Resultado da função

        Raises:
            Exception: Se todas as tentativas falharem
        """
        import pythoncom

        for attempt in range(max_retries):
            try:
                return func()
            except pythoncom.com_error as e:
                if e.hresult == COMErrorHandler.RPC_E_CALL_REJECTED:
                    if attempt < max_retries - 1:
                        # Delay crescente: 0.5s, 1s, 1.5s
                        delay = 0.5 * (attempt + 1)
                        time.sleep(delay)
                        continue
                raise
            except Exception as e:
                raise
        raise Exception(f"{operation_name} falhou após {max_retries} tentativas")


class AutocadWorker(QThread):
    progress = Signal(int)
    finished = Signal(dict)
    error = Signal(str)
    log = Signal(str)
    current_file = Signal(str)
    cancelled = Signal()

    def __init__(self, excel_path, template_folder, fast_mode=False):
        super().__init__()
        self.excel_path = excel_path
        self.template_folder = template_folder
        self.fast_mode = fast_mode
        self._is_cancelled = False

    def cancel_processing(self):
        """Cancela o processamento atual."""
        self._is_cancelled = True

    def run(self):
        try:
            # Estatísticas para o relatório final
            stats = ProcessingStats()

            # Inicializa COM na thread atual
            pythoncom.CoInitialize()

            # Lê o arquivo Excel
            self.log.emit("Lendo arquivo Excel...")
            df = pd.read_excel(self.excel_path)

            # Renomeia coluna 'Name' para 'TipoSuporte' se existir
            if 'Name' in df.columns:
                df = df.rename(columns={'Name': 'TipoSuporte'})
                self.log.emit("Coluna 'Name' renomeada para 'TipoSuporte'")

            # Verifica se todas as colunas necessárias existem
            missing_columns = [col for col in ProcessingConfig.REQUIRED_COLUMNS if col not in df.columns]

            if missing_columns:
                self.error.emit(f"Colunas faltando no Excel: {', '.join(missing_columns)}")
                stats.errors += 1
                stats.error_details.append(f"Colunas faltando: {', '.join(missing_columns)}")
                self.finished.emit(stats.to_dict())
                return

            # Inicializa o AutoCAD
            self.log.emit("Conectando ao AutoCAD...")
            try:
                acad = win32com.client.Dispatch("AutoCAD.Application")
                acad.Visible = False
                self.log.emit("Conexão com AutoCAD estabelecida com sucesso.")
            except Exception as e:
                self.error.emit(f"Erro ao conectar com AutoCAD: {str(e)}")
                stats.errors += 1
                stats.error_details.append(f"Falha na conexão com AutoCAD: {str(e)}")
                self.finished.emit(stats.to_dict())
                return

            # Processa cada linha
            total_rows = len(df)
            stats.total = total_rows

            # Rastreia posições já processadas para detectar duplicatas
            position_counter = {}

            # AGRUPA POR TIPO DE SUPORTE para processamento em lote
            grouped = df.groupby('TipoSuporte')
            total_groups = len(grouped)
            processed_count = 0

            self.log.emit(f"Iniciando processamento de {total_rows} registros em {total_groups} grupo(s) de templates.")

            # Define sleep times baseado no modo
            if self.fast_mode:
                base_sleep_open, base_sleep_save, base_sleep_between = 0.3, 0.05, 0.02
            else:
                base_sleep_open, base_sleep_save, base_sleep_between = 1.0, 0.3, 0.15

            for tipo_suporte, group_df in grouped:
                if self._is_cancelled:
                    self.log.emit("\n⚠️ Processamento cancelado pelo usuário.")
                    self.cancelled.emit()
                    return

                template_path = os.path.join(self.template_folder, f"{tipo_suporte}.dwg")

                # Verifica se o template existe
                if not os.path.exists(template_path):
                    self.log.emit(f"⚠️ Template {tipo_suporte} não encontrado. Pulando {len(group_df)} registros.")
                    stats.template_not_found += len(group_df)
                    for _, row in group_df.iterrows():
                        posicao = str(row['POSICAO'])
                        stats.not_found_details.append(f"{posicao} (Tipo: {tipo_suporte})")
                    processed_count += len(group_df)
                    continue

                self.log.emit(f"\n{'='*50}")
                self.log.emit(f"TEMPLATE: {tipo_suporte}.dwg ({len(group_df)} documentos)")
                self.log.emit(f"{'='*50}")

                # Abre o template para processar este lote
                doc = None
                template_doc = None
                try:
                    template_doc = COMErrorHandler.execute_with_retry(
                        lambda: acad.Documents.Open(template_path),
                        operation_name=f"Open {template_path}",
                        max_retries=2
                    )
                    if template_doc is None:
                        raise Exception("Falha ao abrir template")
                    time.sleep(base_sleep_open)

                    # Processa cada documento deste tipo
                    for idx, (_, row) in enumerate(group_df.iterrows()):
                        i = processed_count + idx
                        if self._is_cancelled:
                            try:
                                template_doc.Close(False)
                            except:
                                pass
                            self.cancelled.emit()
                            return

                        posicao = str(row['POSICAO'])
                        elevacao = str(row['Elevacao']).replace(',', '.')
                        h = str(row['MEDIDA_H']) if pd.notna(row['MEDIDA_H']) else "-"
                        l = str(row['MEDIDA_L']) if pd.notna(row['MEDIDA_L']) else "-"
                        m = str(row['MEDIDA_M']) if pd.notna(row['MEDIDA_M']) else "-"
                        h1 = str(row['MEDIDA_H1']) if pd.notna(row['MEDIDA_H1']) else "-"
                        h2 = str(row['MEDIDA_H2']) if pd.notna(row['MEDIDA_H2']) else "-"
                        l1 = str(row['MEDIDA_L1']) if pd.notna(row['MEDIDA_L1']) else "-"
                        l2 = str(row['MEDIDA_L2']) if pd.notna(row['MEDIDA_L2']) else "-"
                        b = str(row['MEDIDA_B']) if pd.notna(row['MEDIDA_B']) else "-"

                        # Tratamento de duplicatas
                        if posicao not in position_counter:
                            position_counter[posicao] = 1
                            filename_suffix = ""
                        else:
                            position_counter[posicao] += 1
                            filename_suffix = f"_{position_counter[posicao]}"
                            stats.duplicates += 1
                            stats.duplicate_details.append(f"{posicao} -> {posicao}{filename_suffix}")

                        output_filename = f"{posicao}{filename_suffix}.dwg"
                        output_path = os.path.join(os.path.dirname(self.excel_path), output_filename)

                        progress_percent = int((processed_count + idx + 1) / total_rows * 100)
                        self.progress.emit(progress_percent)
                        self.current_file.emit(f"[{idx+1}/{len(group_df)}] {posicao} ({tipo_suporte})")
                        self.log.emit(f"[{i+1}/{total_rows}] {posicao} -> {output_filename}")

                        # Mapeamento de atributos
                        attribute_mapping = {
                            "POSICAO": posicao,
                            "TIPOSUPORTE": tipo_suporte,
                            "ELEVACAO": elevacao,
                            "H": h, "L": l, "M": m,
                            "H1": h1, "H2": h2,
                            "L1": l1, "L2": l2,
                            "B": b,
                            "DATA_ATUAL": datetime.now().strftime('%d/%m/%Y')
                        }

                        # Para documentos após o primeiro, reabre o template
                        if idx > 0:
                            time.sleep(base_sleep_between)
                            doc = COMErrorHandler.execute_with_retry(
                                lambda: acad.Documents.Open(template_path),
                                operation_name=f"Reopen {template_path}",
                                max_retries=2
                            )
                            time.sleep(base_sleep_open * 0.5)
                        else:
                            doc = template_doc
                            time.sleep(base_sleep_between)

                        # Preenche atributos
                        attr_count = 0
                        found_attributes = False

                        def fill_attributes():
                            nonlocal attr_count, found_attributes
                            for entity in doc.PaperSpace:
                                if entity.ObjectName == "AcDbBlockReference" and entity.HasAttributes:
                                    found_attributes = True
                                    for attrib in entity.GetAttributes():
                                        tag = attrib.TagString.upper()
                                        if tag in attribute_mapping:
                                            attrib.TextString = attribute_mapping[tag]
                                            attr_count += 1

                        try:
                            COMErrorHandler.execute_with_retry(
                                fill_attributes,
                                operation_name="Fill attributes",
                                max_retries=2
                            )
                        except:
                            pass

                        if not found_attributes or attr_count == 0:
                            self.log.emit(f"  ⚠️ Sem atributos -> pulado")
                            stats.no_attributes += 1
                            stats.no_attributes_details.append(f"{posicao} (Tipo: {tipo_suporte})")
                            if idx > 0:
                                try:
                                    doc.Close(False)
                                except:
                                    pass
                            doc = None
                            continue

                        # Salva
                        time.sleep(base_sleep_between)
                        COMErrorHandler.execute_with_retry(
                            lambda: doc.SaveAs(output_path),
                            operation_name=f"SaveAs {output_path}",
                            max_retries=2
                        )
                        time.sleep(base_sleep_save)
                        doc.Close()
                        doc = None

                        self.log.emit(f"  ✅ {output_filename} criado!")
                        stats.success += 1

                    # Fecha o template do lote
                    try:
                        template_doc.Close(False)
                    except:
                        pass

                    processed_count += len(group_df)

                except Exception as e:
                    self.log.emit(f"❌ Erro no grupo {tipo_suporte}: {str(e)}")
                    stats.errors += len(group_df)
                    for _, row in group_df.iterrows():
                        stats.error_details.append(f"{str(row['POSICAO'])}: {str(e)}")
                    try:
                        if doc:
                            doc.Close(False)
                        try:
                            template_doc.Close(False)
                        except:
                            pass
                    except:
                        pass

            self.log.emit("\n===== PROCESSAMENTO CONCLUÍDO =====")
            self.finished.emit(stats.to_dict())

        except Exception as e:
            self.error.emit(f"Erro geral: {str(e)}")
            stats = ProcessingStats()
            stats.errors = 1
            stats.error_details.append(f"Erro geral: {str(e)}")
            self.finished.emit(stats.to_dict())
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

        # Botão de cancelar (inicialmente oculto)
        self.cancel_button = QPushButton("Cancelar Processamento")
        self.cancel_button.clicked.connect(self.cancel_processing)
        self.cancel_button.setEnabled(False)
        self.cancel_button.setStyleSheet("background-color: #ffcccc;")

        # Checkbox de modo rápido
        self.fast_mode_checkbox = QCheckBox("Modo Rápido ⚡")
        self.fast_mode_checkbox.setToolTip("Reduz tempo de processamento (use apenas se estável)")
        self.fast_mode_checkbox.setChecked(False)

        # Área de progresso com detalhes
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_label = QLabel("Aguardando início...")
        self.progress_label.setAlignment(Qt.AlignCenter)
        progress_vbox = QVBoxLayout()
        progress_vbox.addWidget(self.progress_label)
        progress_vbox.addWidget(self.progress_bar)
        progress_layout.addLayout(progress_vbox, 1)

        # Adiciona widgets ao layout de controles
        controls_layout.addLayout(excel_layout)
        controls_layout.addLayout(template_layout)
        controls_layout.addWidget(self.process_button)
        controls_layout.addWidget(self.cancel_button)
        controls_layout.addWidget(self.fast_mode_checkbox)
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
        self.worker = AutocadWorker(
            self.excel_path,
            self.template_folder,
            fast_mode=self.fast_mode_checkbox.isChecked()
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.log.connect(self.add_to_log)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(self.processing_finished)
        self.worker.cancelled.connect(self.processing_cancelled)
        self.worker.current_file.connect(self.update_current_file)
        self.worker.start()

        # Mostrar botão de cancelar
        self.cancel_button.setEnabled(True)

    def cancel_processing(self):
        """Solicita o cancelamento do processamento."""
        reply = QMessageBox.question(
            self,
            "Confirmar Cancelamento",
            "Deseja realmente cancelar o processamento?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.worker.cancel_processing()
            self.cancel_button.setEnabled(False)

    def processing_cancelled(self):
        """Chamado quando o processamento é cancelado."""
        self.add_to_log("\n" + "=" * 50)
        self.add_to_log("PROCESSAMENTO CANCELADO PELO USUÁRIO")
        self.add_to_log("=" * 50)
        self.cancel_button.setEnabled(False)
        self.excel_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.progress_label.setText("Processamento cancelado")

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_current_file(self, file_info):
        """Atualiza o label com o arquivo atual."""
        self.progress_label.setText(file_info)

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
        self.add_to_log(f"Arquivos criados com sucesso: {stats['success']}")
        self.add_to_log(f"Templates não encontrados: {stats['template_not_found']}")
        self.add_to_log(f"Templates sem atributos: {stats.get('no_attributes', 0)}")
        self.add_to_log(f"Posicoes duplicatas tratadas: {stats.get('duplicates', 0)}")
        self.add_to_log(f"Erros durante o processamento: {stats['errors']}")

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

        # Detalhes sobre posicoes duplicatas
        if stats.get("duplicates", 0) > 0:
            self.add_to_log("\nDetalhes de Posicoes Duplicatas:")
            for detail in stats["duplicate_details"]:
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
        self.cancel_button.setEnabled(False)
        self.progress_label.setText("Processamento concluído")

        # Mostra resumo em uma mensagem
        QMessageBox.information(
            self,
            "Processamento Concluido",
            f"Processamento concluido!\n\n"
            f"Total: {stats['total']}\n"
            f"Sucesso: {stats['success']}\n"
            f"Templates nao encontrados: {stats['template_not_found']}\n"
            f"Templates sem atributos: {stats.get('no_attributes', 0)}\n"
            f"Posicoes duplicatas tratadas: {stats.get('duplicates', 0)}\n"
            f"Erros: {stats['errors']}\n\n"
            f"Veja o log para mais detalhes."
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
