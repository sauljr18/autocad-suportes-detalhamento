import os
import sys
from datetime import datetime

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (QApplication, QCheckBox, QFileDialog, QHBoxLayout,
                               QLabel, QMainWindow, QMessageBox, QProgressBar,
                               QPushButton, QSplitter, QTextEdit, QVBoxLayout,
                               QWidget)

import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib
matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt


class ProcessingConfig:
    """Configurações centralizadas do processamento DXF."""

    # Colunas obrigatórias do Excel
    REQUIRED_COLUMNS = ['POSICAO', 'TipoSuporte', 'Elevacao',
                        'MEDIDA_H', 'MEDIDA_L', 'MEDIDA_M',
                        'MEDIDA_H1', 'MEDIDA_H2', 'MEDIDA_L1',
                        'MEDIDA_L2', 'MEDIDA_B']

    # Mapeamento de atributos do DXF
    ATTRIBUTE_TAGS = ["POSICAO", "TIPOSUPORTE", "ELEVACAO",
                      "H", "L", "M", "H1", "H2", "L1", "L2", "B", "DATA_ATUAL"]

    # Extensões de arquivo
    TEMPLATE_EXTENSION = ".dxf"
    OUTPUT_EXTENSION = ".dxf"


class ProcessingStats:
    """Gerencia estatísticas do processamento."""

    def __init__(self):
        self.total = 0
        self.success = 0
        self.template_not_found = 0
        self.errors = 0
        self.no_attributes = 0
        self.duplicates = 0
        self.pdf_generated = 0
        self.pdf_failed = 0
        self.error_details = []
        self.not_found_details = []
        self.no_attributes_details = []
        self.duplicate_details = []
        self.pdf_failed_details = []

    def to_dict(self):
        """Converte para dicionário para sinal."""
        return {
            "total": self.total,
            "success": self.success,
            "template_not_found": self.template_not_found,
            "errors": self.errors,
            "no_attributes": self.no_attributes,
            "duplicates": self.duplicates,
            "pdf_generated": self.pdf_generated,
            "pdf_failed": self.pdf_failed,
            "error_details": self.error_details,
            "not_found_details": self.not_found_details,
            "no_attributes_details": self.no_attributes_details,
            "duplicate_details": self.duplicate_details,
            "pdf_failed_details": self.pdf_failed_details
        }


class DXFWorker(QThread):
    """Worker thread para processamento DXF com ezdxf."""

    progress = Signal(int)
    finished = Signal(dict)
    error = Signal(str)
    log = Signal(str)
    current_file = Signal(str)
    cancelled = Signal()

    def __init__(self, excel_path, template_folder, generate_pdf=False):
        super().__init__()
        self.excel_path = excel_path
        self.template_folder = template_folder
        self.generate_pdf = generate_pdf
        self._is_cancelled = False

    def cancel_processing(self):
        """Cancela o processamento atual."""
        self._is_cancelled = True

    def process_document(self, template_path, output_path, attribute_mapping):
        """
        Processa um documento DXF - lê template, modifica atributos, salva.

        Args:
            template_path: Caminho do template DXF
            output_path: Caminho de saída do DXF modificado
            attribute_mapping: Dicionário {tag: valor} para preencher atributos

        Returns:
            Tuple (success, attr_count, error_message)
        """
        try:
            # Lê o template DXF
            doc = ezdxf.readfile(template_path)

            attr_count = 0
            found_attributes = False

            # Processa todos os layouts (itera diretamente sobre Layout objects)
            for layout in doc.layouts:
                layout_name = layout.name
                if layout_name == "Model":
                    continue

                # Busca entidades INSERT (block references) com atributos
                for entity in layout:
                    if entity.dxftype() == 'INSERT':
                        try:
                            for attrib in entity.attribs:
                                found_attributes = True
                                # Converte tag para string antes de upper()
                                tag_value = attrib.dxf.tag
                                if tag_value is not None:
                                    tag = str(tag_value).upper()
                                    if tag in attribute_mapping:
                                        attrib.dxf.text = attribute_mapping[tag]
                                        attr_count += 1
                        except Exception:
                            # Ignora erros em atributos individuais
                            pass

            # Verifica se encontrou e modificou atributos
            if not found_attributes or attr_count == 0:
                return False, 0, "Sem atributos encontrados"

            # Salva o documento modificado
            doc.saveas(output_path)
            return True, attr_count, None

        except Exception as e:
            return False, 0, str(e)

    def convert_to_pdf(self, dxf_path, pdf_path):
        """
        Converte arquivo DXF para PDF usando matplotlib.

        Returns: (success, error_message)
        """
        try:
            doc = ezdxf.readfile(dxf_path)
            ctx = RenderContext(doc)

            # Cria figura matplotlib
            fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 size

            # Renderiza apenas PaperSpace layouts
            layout_found = False
            for layout in doc.layouts:
                if layout.name == "Model":
                    continue
                Frontend(ctx, MatplotlibBackend(ax)).draw_layout(layout)
                layout_found = True
                break  # Primeiro PaperSpace apenas

            if not layout_found:
                plt.close(fig)
                return False, "Nenhum layout PaperSpace encontrado"

            # Salva como PDF
            fig.savefig(pdf_path, format='pdf', bbox_inches='tight', dpi=300)
            plt.close(fig)
            return True, None

        except Exception as e:
            return False, str(e)

    def run(self):
        """Executa o processamento dos dados."""
        try:
            stats = ProcessingStats()

            # Lê o arquivo Excel
            self.log.emit("Lendo arquivo Excel...")
            df = pd.read_excel(self.excel_path)

            # Renomeia coluna 'Name' para 'TipoSuporte' se existir
            if 'Name' in df.columns:
                df = df.rename(columns={'Name': 'TipoSuporte'})
                self.log.emit("Coluna 'Name' renomeada para 'TipoSuporte'")

            # Verifica se todas as colunas necessárias existem
            missing_columns = [col for col in ProcessingConfig.REQUIRED_COLUMNS
                             if col not in df.columns]

            if missing_columns:
                self.error.emit(f"Colunas faltando: {', '.join(missing_columns)}")
                stats.errors = 1
                stats.error_details.append(f"Colunas faltando: {', '.join(missing_columns)}")
                self.finished.emit(stats.to_dict())
                return

            # Agrupa por TipoSuporte para processamento eficiente
            grouped = df.groupby('TipoSuporte')
            total_rows = len(df)
            stats.total = total_rows
            processed_count = 0

            self.log.emit(f"Processando {total_rows} registros em {len(grouped)} grupo(s).")

            # Rastreia posições já processadas para detectar duplicatas
            position_counter = {}

            for tipo_suporte, group_df in grouped:
                if self._is_cancelled:
                    self.log.emit("\n⚠️ Processamento cancelado pelo usuário.")
                    self.cancelled.emit()
                    return

                template_path = os.path.join(
                    self.template_folder,
                    f"{tipo_suporte}.dxf"
                )

                # Verifica se o template existe
                if not os.path.exists(template_path):
                    self.log.emit(f"⚠️ Template {tipo_suporte}.dxf não encontrado.")
                    stats.template_not_found += len(group_df)
                    for _, row in group_df.iterrows():
                        posicao = str(row['POSICAO'])
                        stats.not_found_details.append(
                            f"{posicao} (Tipo: {tipo_suporte})"
                        )
                    processed_count += len(group_df)
                    continue

                self.log.emit(f"\n{'='*50}")
                self.log.emit(f"TEMPLATE: {tipo_suporte}.dxf ({len(group_df)} docs)")
                self.log.emit(f"{'='*50}")

                # Processa cada documento deste tipo
                for idx, (_, row) in enumerate(group_df.iterrows()):
                    if self._is_cancelled:
                        self.cancelled.emit()
                        return

                    i = processed_count + idx

                    # Extrai dados do Excel
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
                        stats.duplicate_details.append(
                            f"{posicao} -> {posicao}{filename_suffix}"
                        )

                    output_filename = f"{posicao}{filename_suffix}.dxf"
                    output_path = os.path.join(
                        os.path.dirname(self.excel_path),
                        output_filename
                    )

                    progress_percent = int((i + 1) / total_rows * 100)
                    self.progress.emit(progress_percent)
                    self.current_file.emit(f"[{idx+1}/{len(group_df)}] {posicao}")

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

                    # Processa documento (sem delays!)
                    success, attr_count, error_msg = self.process_document(
                        template_path, output_path, attribute_mapping
                    )

                    if success:
                        self.log.emit(f"[{i+1}/{total_rows}] ✅ {output_filename} ({attr_count} atribs)")
                        stats.success += 1

                        # Gera PDF se habilitado
                        if self.generate_pdf:
                            pdf_folder = os.path.join(os.path.dirname(self.excel_path), "Pdf")
                            os.makedirs(pdf_folder, exist_ok=True)

                            pdf_filename = f"{posicao}{filename_suffix}.pdf"
                            pdf_path = os.path.join(pdf_folder, pdf_filename)

                            pdf_success, pdf_error = self.convert_to_pdf(output_path, pdf_path)
                            if pdf_success:
                                self.log.emit(f"      📄 {pdf_filename} criado")
                                stats.pdf_generated += 1
                            else:
                                self.log.emit(f"      ⚠️ PDF falhou: {pdf_error}")
                                stats.pdf_failed += 1
                                stats.pdf_failed_details.append(
                                    f"{posicao}: {pdf_error}"
                                )
                    else:
                        if error_msg == "Sem atributos encontrados":
                            self.log.emit(f"  ⚠️ Sem atributos")
                            stats.no_attributes += 1
                            stats.no_attributes_details.append(
                                f"{posicao} (Tipo: {tipo_suporte})"
                            )
                        else:
                            self.log.emit(f"  ❌ Erro: {error_msg}")
                            stats.errors += 1
                            stats.error_details.append(f"{posicao}: {error_msg}")

                processed_count += len(group_df)

            self.log.emit("\n===== PROCESSAMENTO CONCLUÍDO =====")
            self.finished.emit(stats.to_dict())

        except Exception as e:
            self.error.emit(f"Erro geral: {str(e)}")
            stats = ProcessingStats()
            stats.errors = 1
            stats.error_details.append(f"Erro geral: {str(e)}")
            self.finished.emit(stats.to_dict())


class MainWindow(QMainWindow):
    """Janela principal da aplicação de processamento DXF."""

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Integração Excel-DXF (Multi-plataforma)")
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

        # Checkbox para gerar PDF
        self.pdf_checkbox = QCheckBox("Gerar Pdf's")
        self.pdf_checkbox.setToolTip("Gerar arquivos PDF junto com DXF")
        self.pdf_checkbox.setChecked(False)

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
        controls_layout.addWidget(self.pdf_checkbox)
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
        """Abre diálogo para selecionar arquivo Excel."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx *.xls)"
        )

        if file_path:
            self.excel_path = file_path
            self.excel_label.setText(f"Excel: {os.path.basename(file_path)}")
            self.check_ready()

    def select_template_folder(self):
        """Abre diálogo para selecionar pasta de templates."""
        folder_path = QFileDialog.getExistingDirectory(
            self, "Selecionar Pasta de Templates"
        )

        if folder_path:
            self.template_folder = folder_path
            self.template_label.setText(f"Templates: {os.path.basename(folder_path)}")
            self.check_ready()

    def check_ready(self):
        """Verifica se ambos os arquivos foram selecionados."""
        if self.excel_path and self.template_folder:
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)

    def process_data(self):
        """Inicia o processamento dos dados."""
        # Limpa o log
        self.log_text.clear()

        # Desativa os botões durante o processamento
        self.excel_button.setEnabled(False)
        self.template_button.setEnabled(False)
        self.process_button.setEnabled(False)

        # Adiciona cabeçalho ao log
        from datetime import datetime as dt
        self.add_to_log(f"=== INÍCIO DO PROCESSAMENTO ===")
        self.add_to_log(f"Data/Hora: {dt.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.add_to_log(f"Arquivo Excel: {self.excel_path}")
        self.add_to_log(f"Pasta de Templates: {self.template_folder}")
        self.add_to_log("-" * 50)

        # Cria e inicia o worker thread
        self.worker = DXFWorker(
            self.excel_path,
            self.template_folder,
            generate_pdf=self.pdf_checkbox.isChecked()
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
        """Atualiza a barra de progresso."""
        self.progress_bar.setValue(value)

    def update_current_file(self, file_info):
        """Atualiza o label com o arquivo atual."""
        self.progress_label.setText(file_info)

    def add_to_log(self, message):
        """Adiciona mensagem ao log."""
        self.log_text.append(message)
        # Rola para o fim do texto
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def show_error(self, message):
        """Exibe mensagem de erro."""
        QMessageBox.critical(self, "Erro", message)
        self.add_to_log(f"❌ ERRO: {message}")

    def processing_finished(self, stats):
        """Chamado ao final do processamento."""
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

        # Estatísticas de PDF
        if stats.get('pdf_generated', 0) > 0 or stats.get('pdf_failed', 0) > 0:
            self.add_to_log(f"PDFs gerados com sucesso: {stats.get('pdf_generated', 0)}")
            self.add_to_log(f"PDFs falhados: {stats.get('pdf_failed', 0)}")

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

        # Detalhes sobre PDFs falhados
        if stats.get("pdf_failed", 0) > 0:
            self.add_to_log("\nDetalhes de PDFs Falhados:")
            for detail in stats.get("pdf_failed_details", []):
                self.add_to_log(f"  - {detail}")

        self.add_to_log("\n" + "=" * 50)
        from datetime import datetime as dt
        self.add_to_log(f"Processamento finalizado em: {dt.now().strftime('%d/%m/%Y %H:%M:%S')}")

        # Reativa os botões
        self.excel_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.progress_label.setText("Processamento concluído")

        # Mostra resumo em uma mensagem
        summary = (
            f"Processamento concluido!\n\n"
            f"Total: {stats['total']}\n"
            f"Sucesso: {stats['success']}\n"
            f"Templates nao encontrados: {stats['template_not_found']}\n"
            f"Templates sem atributos: {stats.get('no_attributes', 0)}\n"
            f"Posicoes duplicatas tratadas: {stats.get('duplicates', 0)}\n"
            f"Erros: {stats['errors']}"
        )
        if stats.get('pdf_generated', 0) > 0 or stats.get('pdf_failed', 0) > 0:
            summary += (
                f"\n\nPDFs:\n"
                f"  Gerados: {stats.get('pdf_generated', 0)}\n"
                f"  Falhados: {stats.get('pdf_failed', 0)}"
            )
        summary += "\n\nVeja o log para mais detalhes."

        QMessageBox.information(
            self,
            "Processamento Concluido",
            summary
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
