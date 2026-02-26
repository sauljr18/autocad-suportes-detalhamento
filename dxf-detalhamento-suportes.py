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
                        'MEDIDA_L2', 'MEDIDA_B',
                        'NUM_DOC', 'QTD', 'CLIENTE']

    # Mapeamento de atributos do DXF
    ATTRIBUTE_TAGS = ["POSICAO", "TIPOSUPORTE", "ELEVACAO",
                      "H", "L", "M", "H1", "H2", "L1", "L2", "B", "DATA_ATUAL",
                      "NUM_DOC", "QTD", "CLIENTE"]

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


class DXFConversionWorker(QThread):
    """Worker thread para conversão DWG -> DXF em lote."""

    progress = Signal(int)
    finished = Signal(dict)
    error = Signal(str)
    log = Signal(str)
    current_file = Signal(str)
    cancelled = Signal()

    def __init__(self, source_folder, dxf_version="R2013"):
        super().__init__()
        self.source_folder = source_folder
        self.dxf_version = dxf_version
        self._is_cancelled = False

    def cancel_processing(self):
        """Cancela a conversao."""
        self._is_cancelled = True

    def run(self):
        """Executa a conversao dos arquivos DWG para DXF."""
        try:
            import pythoncom
            import win32com.client
            import time

            stats = {
                'total': 0,
                'success': 0,
                'errors': 0,
                'skipped': 0,
                'error_details': []
            }

            # Inicializa COM
            pythoncom.CoInitialize()
            self.log.emit("Conectando ao AutoCAD...")

            # Conecta ao AutoCAD
            try:
                acad = win32com.client.Dispatch("AutoCAD.Application")
                acad.Visible = False
                self.log.emit("Conexao com AutoCAD estabelecida.")
            except Exception as e:
                self.error.emit(f"Erro ao conectar com AutoCAD: {str(e)}")
                self.finished.emit(stats)
                return

            # Busca arquivos DWG (case-insensitive para pegar .dwg e .DWG)
            import glob
            dwg_pattern = os.path.join(self.source_folder, "*.dwg")
            dwg_files = glob.glob(dwg_pattern)

            # Tenta também com maiúscula se não encontrar
            if not dwg_files:
                dwg_pattern_upper = os.path.join(self.source_folder, "*.DWG")
                dwg_files = glob.glob(dwg_pattern_upper)

            # Debug: mostra o que está procurando
            self.log.emit(f"Pasta pesquisada: {self.source_folder}")
            self.log.emit(f"Padrao de busca: {dwg_pattern}")

            if not dwg_files:
                self.log.emit("Nenhum arquivo .dwg encontrado na pasta.")
                # Lista todos os arquivos da pasta para debug
                try:
                    all_files = os.listdir(self.source_folder)
                    self.log.emit(f"Arquivos na pasta: {len(all_files)} itens")
                    # Mostra primeiros 10 arquivos
                    for f in all_files[:10]:
                        self.log.emit(f"  - {f}")
                    if len(all_files) > 10:
                        self.log.emit(f"  ... e mais {len(all_files) - 10} arquivos")
                except Exception as e:
                    self.log.emit(f"Erro ao listar pasta: {e}")
                self.finished.emit(stats)
                return

            stats['total'] = len(dwg_files)
            self.log.emit(f"Encontrados: {len(dwg_files)} arquivos .dwg")

            # Processa cada arquivo
            for i, dwg_path in enumerate(dwg_files):
                if self._is_cancelled:
                    self.log.emit("Conversao cancelada pelo usuario.")
                    self.cancelled.emit()
                    return

                dwg_filename = os.path.basename(dwg_path)
                dxf_path = os.path.splitext(dwg_path)[0] + ".dxf"

                # Verifica se DXF ja existe e eh mais recente
                if os.path.exists(dxf_path):
                    dwg_mtime = os.path.getmtime(dwg_path)
                    dxf_mtime = os.path.getmtime(dxf_path)
                    if dxf_mtime > dwg_mtime:
                        self.log.emit(f"[{i+1}/{len(dwg_files)}] Pulado: {dwg_filename} -> DXF ja atual")
                        stats['skipped'] += 1
                        self.progress.emit(int((i + 1) / len(dwg_files) * 100))
                        continue

                self.current_file.emit(f"[{i+1}/{len(dwg_files)}] {dwg_filename}")
                self.progress.emit(int((i + 1) / len(dwg_files) * 100))

                # Sistema de retry: ate 3 tentativas
                max_retries = 3
                success = False
                last_error = None

                for retry in range(max_retries):
                    if self._is_cancelled:
                        self.cancelled.emit()
                        return

                    try:
                        # Converte caminho para formato do AutoCAD (barra invertida)
                        dwg_path_acad = os.path.abspath(dwg_path).replace("/", "\\")
                        dxf_path_acad = os.path.abspath(dxf_path).replace("/", "\\")

                        # Salva tamanho atual do arquivo DXF (se existir)
                        old_size = 0
                        if os.path.exists(dxf_path):
                            old_size = os.path.getsize(dxf_path)

                        # Abre o DWG
                        doc = acad.Documents.Open(dwg_path_acad)
                        time.sleep(0.8)

                        # Exporta para DXF usando DXFOUT via SendCommand
                        # Codigo da versao DXF: 12 = R2007, 14 = R2010, 16 = R2013, 18 = R2018
                        version_code = "16"  # R2013

                        # Envia comando DXFOUT
                        cmd = f'(command "DXFOUT" "{dxf_path_acad}" "" "{version_code}") '
                        doc.SendCommand(cmd)

                        # Aguarda o arquivo DXF ser criado/atualizado
                        max_wait = 10  # segundos maximos de espera
                        wait_count = 0
                        file_created = False

                        while wait_count < max_wait:
                            time.sleep(0.5)
                            wait_count += 0.5

                            if os.path.exists(dxf_path):
                                new_size = os.path.getsize(dxf_path)
                                # Verifica se o arquivo cresceu (ou seja, foi reescrito)
                                if new_size > old_size or new_size > 1000:  # DXF valido tem pelo menos 1KB
                                    file_created = True
                                    break

                        # Fecha sem salvar
                        doc.Close(False)
                        time.sleep(0.5)  # Delay maior entre tentativas

                        # Verifica se o arquivo DXF foi realmente criado
                        if os.path.exists(dxf_path) and os.path.getsize(dxf_path) > 1000:
                            if retry > 0:
                                self.log.emit(f"[{i+1}/{len(dwg_files)}] Sucesso (tentativa {retry + 1}): {dwg_filename} -> DXF")
                            else:
                                self.log.emit(f"[{i+1}/{len(dwg_files)}] Sucesso: {dwg_filename} -> DXF")
                            stats['success'] += 1
                            success = True
                            break
                        else:
                            raise Exception("Arquivo DXF nao foi criado ou esta vazio")

                    except Exception as e:
                        last_error = str(e)
                        # Tenta fechar o documento se estiver aberto
                        try:
                            if acad.Documents.Count > 0:
                                acad.ActiveDocument.Close(False)
                        except:
                            pass

                        # Se nao for a ultima tentativa, aguarda antes de tentar novamente
                        if retry < max_retries - 1:
                            time.sleep(1.0)  # Espera 1 segundo antes de tentar de novo

                # Se apos todas as tentativas ainda falhou
                if not success:
                    stats['errors'] += 1
                    stats['error_details'].append(f"{dwg_filename}: {last_error}")
                    self.log.emit(f"[{i+1}/{len(dwg_files)}] Erro (apos {max_retries} tentativas): {dwg_filename}: {last_error}")

            self.log.emit("\n===== CONVERSAO CONCLUIDA =====")
            self.finished.emit(stats)

        except Exception as e:
            self.error.emit(f"Erro geral: {str(e)}")
            self.finished.emit(stats)


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

                    # Extrai novos campos do Excel (tratando valores NaN)
                    num_doc = str(row['NUM_DOC']) if pd.notna(row['NUM_DOC']) else ""
                    qtd = str(row['QTD']) if pd.notna(row['QTD']) else ""
                    cliente = str(row['CLIENTE']) if pd.notna(row['CLIENTE']) else ""

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
                        "DATA_ATUAL": datetime.now().strftime('%d/%m/%Y'),
                        # Novos atributos
                        "NUM_DOC": num_doc,
                        "QTD": qtd,
                        "CLIENTE": cliente
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

        # Botao de conversao DWG->DXF
        self.convert_button = QPushButton("Converter DWG->DXF (Lote)")
        self.convert_button.clicked.connect(self.convert_dwg_to_dxf)
        self.convert_button.setStyleSheet("background-color: #e6f3ff;")

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
        controls_layout.addWidget(self.convert_button)
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
        self.conversion_worker = None

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

    def convert_dwg_to_dxf(self):
        """Inicia a conversao de DWG para DXF em lote."""
        # Se pasta de templates ja estiver selecionada, usa ela
        if self.template_folder:
            folder = self.template_folder
        else:
            # Senao, pede para selecionar a pasta
            folder_path = QFileDialog.getExistingDirectory(
                self, "Selecionar Pasta com Arquivos DWG"
            )
            if not folder_path:
                return
            folder = folder_path

        # Confirmacao
        reply = QMessageBox.question(
            self,
            "Confirmar Conversao",
            f"Converter todos os arquivos .dwg da pasta:\n{folder}\n\nPara formato DXF (versao 2013)?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # Limpa o log
        self.log_text.clear()

        # Desativa botoes durante conversao
        self.excel_button.setEnabled(False)
        self.template_button.setEnabled(False)
        self.process_button.setEnabled(False)
        self.convert_button.setEnabled(False)

        # Adiciona cabecalho ao log
        from datetime import datetime as dt
        self.add_to_log(f"=== CONVERSAO DWG -> DXF ===")
        self.add_to_log(f"Data/Hora: {dt.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.add_to_log(f"Pasta: {folder}")
        self.add_to_log("-" * 50)

        # Cria e inicia o worker de conversao
        self.conversion_worker = DXFConversionWorker(folder, dxf_version="R2013")
        self.conversion_worker.progress.connect(self.update_progress)
        self.conversion_worker.log.connect(self.add_to_log)
        self.conversion_worker.error.connect(self.show_error)
        self.conversion_worker.finished.connect(self.conversion_finished)
        self.conversion_worker.cancelled.connect(self.conversion_cancelled)
        self.conversion_worker.current_file.connect(self.update_current_file)
        self.conversion_worker.start()

    def conversion_cancelled(self):
        """Chamado quando a conversao e cancelada."""
        self.add_to_log("\n" + "=" * 50)
        self.add_to_log("CONVERSAO CANCELADA PELO USUARIO")
        self.add_to_log("=" * 50)
        self.excel_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.convert_button.setEnabled(True)
        self.progress_label.setText("Conversao cancelada")

    def conversion_finished(self, stats):
        """Chamado ao final da conversao."""
        self.add_to_log("\n" + "=" * 50)
        self.add_to_log("RELATORIO FINAL DE CONVERSAO")
        self.add_to_log("-" * 50)
        self.add_to_log(f"Total de arquivos: {stats['total']}")
        self.add_to_log(f"Convertidos com sucesso: {stats['success']}")
        self.add_to_log(f"Ja atualizados (pulados): {stats.get('skipped', 0)}")
        self.add_to_log(f"Erros: {stats['errors']}")

        if stats['error_details']:
            self.add_to_log("\nDetalhes de Erros:")
            for detail in stats['error_details']:
                self.add_to_log(f"  - {detail}")

        self.add_to_log("\n" + "=" * 50)
        from datetime import datetime as dt
        self.add_to_log(f"Conversao finalizada em: {dt.now().strftime('%d/%m/%Y %H:%M:%S')}")

        # Reativa os botoes
        self.excel_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.convert_button.setEnabled(True)
        self.progress_label.setText("Conversao concluida")

        # Mostra resumo
        summary = (
            f"Conversao concluida!\n\n"
            f"Total: {stats['total']}\n"
            f"Sucesso: {stats['success']}\n"
            f"Pulados: {stats.get('skipped', 0)}\n"
            f"Erros: {stats['errors']}"
        )
        QMessageBox.information(self, "Conversao Concluida", summary)

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
        # Verifica se ha worker ativo
        has_active_worker = (self.worker and self.worker.isRunning()) or \
                            (self.conversion_worker and self.conversion_worker.isRunning())

        if not has_active_worker:
            return

        reply = QMessageBox.question(
            self,
            "Confirmar Cancelamento",
            "Deseja realmente cancelar o processamento atual?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if self.worker and self.worker.isRunning():
                self.worker.cancel_processing()
            if self.conversion_worker and self.conversion_worker.isRunning():
                self.conversion_worker.cancel_processing()
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
