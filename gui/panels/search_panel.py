"""Painel de busca e filtros."""

from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QToolButton, QVBoxLayout, QWidget, QMenu,
    QInputDialog, QMessageBox, QListWidget, QListWidgetItem,
    QDialog, QDialogButtonBox, QTextEdit
)

from core.models import FiltroBusca


class FiltroItemWidget(QFrame):
    """Widget que representa um filtro ativo."""

    removido = Signal()

    def __init__(self, filtro: FiltroBusca, parent=None):
        super().__init__(parent)
        self._filtro = filtro

        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setStyleSheet("""
            QFrame {
                background-color: #e3f2fd;
                border-radius: 4px;
                padding: 2px;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)

        # Texto do filtro
        label = QLabel(str(filtro))
        label.setStyleSheet("color: #1565c0;")

        # Botão remover
        btn_remover = QToolButton()
        btn_remover.setText("✕")
        btn_remover.setStyleSheet("""
            QToolButton {
                border: none;
                color: #d32f2f;
                font-weight: bold;
            }
            QToolButton:hover {
                color: #b71c1c;
            }
        """)
        btn_remover.clicked.connect(self.removido.emit)

        layout.addWidget(label)
        layout.addWidget(btn_remover)


class SearchPanel(QWidget):
    """
    Painel de busca e filtros.

    Sinais:
        busca_solicitada: Emitido quando busca é solicitada (texto, filtros)
        filtro_adicionado: Emitido quando filtro é adicionado
        filtro_removido: Emitido quando filtro é removido
        preset_carregado: Emitido quando preset é carregado
    """

    busca_solicitada = Signal(str, list)
    filtro_adicionado = Signal(object)
    filtro_removido = Signal(int)
    preset_carregado = Signal(str)
    limpar_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self._filtros_ativos: List[FiltroBusca] = []
        self._filtros_widgets: List[FiltroItemWidget] = []
        self._campos_disponiveis: List[Dict[str, str]] = []

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Configura a UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Campo de busca geral
        busca_layout = QHBoxLayout()

        self._busca_input = QLineEdit()
        self._busca_input.setPlaceholderText("Buscar em todas as colunas...")
        self._busca_input.returnPressed.connect(self._on_buscar)

        self._btn_buscar = QPushButton("🔍 Buscar")
        self._btn_buscar.clicked.connect(self._on_buscar)

        busca_layout.addWidget(self._busca_input)
        busca_layout.addWidget(self._btn_buscar)

        layout.addLayout(busca_layout)

        # Linha de filtros
        filtros_layout = QHBoxLayout()

        # Combo de campo
        self._combo_campo = QComboBox()
        self._combo_campo.setMinimumWidth(150)
        self._combo_campo.currentTextChanged.connect(self._on_campo_changed)

        # Combo de operador
        self._combo_operador = QComboBox()
        self._combo_operador.setMinimumWidth(120)

        # Campo de valor
        self._valor_input = QLineEdit()
        self._valor_input.setMinimumWidth(150)
        self._valor_input.setPlaceholderText("Valor...")

        # Botão adicionar filtro
        self._btn_adicionar_filtro = QPushButton("+ Adicionar Filtro")
        self._btn_adicionar_filtro.clicked.connect(self._on_adicionar_filtro)

        filtros_layout.addWidget(QLabel("Filtrar por:"))
        filtros_layout.addWidget(self._combo_campo)
        filtros_layout.addWidget(self._combo_operador)
        filtros_layout.addWidget(self._valor_input)
        filtros_layout.addWidget(self._btn_adicionar_filtro)
        filtros_layout.addStretch()

        layout.addLayout(filtros_layout)

        # Área de filtros ativos
        self._filtros_container = QWidget()
        self._filtros_layout = QHBoxLayout(self._filtros_container)
        self._filtros_layout.setContentsMargins(0, 0, 0, 0)
        self._filtros_layout.addStretch()

        layout.addWidget(QLabel("Filtros ativos:"))
        layout.addWidget(self._filtros_container)

        # Linha de presets e histórico
        presets_layout = QHBoxLayout()

        # Presets
        self._btn_presets = QPushButton("📁 Presets")
        self._btn_presets.clicked.connect(self._mostrar_menu_presets)

        self._btn_salvar_preset = QPushButton("💾 Salvar Preset")
        self._btn_salvar_preset.clicked.connect(self._salvar_preset)

        # Histórico
        self._btn_historico_anterior = QPushButton("◀ Anterior")
        self._btn_historico_anterior.clicked.connect(lambda: self.historico_navegar.emit(-1))

        self._btn_historico_proximo = QPushButton("Próximo ▶")
        self._btn_historico_proximo.clicked.connect(lambda: self.historico_navegar.emit(1))

        self._btn_limpar = QPushButton("🗑 Limpar")
        self._btn_limpar.clicked.connect(self._on_limpar)

        presets_layout.addWidget(self._btn_presets)
        presets_layout.addWidget(self._btn_salvar_preset)
        presets_layout.addStretch()
        presets_layout.addWidget(self._btn_historico_anterior)
        presets_layout.addWidget(self._btn_historico_proximo)
        presets_layout.addWidget(self._btn_limpar)

        layout.addLayout(presets_layout)

        # Sinal adicional para navegação no histórico
        self.historico_navegar = Signal(int)

    def definir_campos(self, campos: List[Dict[str, str]]) -> None:
        """
        Define os campos disponíveis para filtro.

        Args:
            campos: Lista de dicts com 'nome', 'tipo', 'label'
        """
        self._campos_disponiveis = campos

        self._combo_campo.clear()
        for campo in campos:
            self._combo_campo.addItem(campo['label'], campo['nome'])

        self._on_campo_changed(self._combo_campo.currentText())

    def _on_campo_changed(self, label: str) -> None:
        """Atualiza operadores baseado no campo selecionado."""
        nome_campo = self._combo_campo.currentData()

        # Encontra o tipo do campo
        tipo = 'texto'
        for campo in self._campos_disponiveis:
            if campo['nome'] == nome_campo:
                tipo = campo['tipo']
                break

        self._combo_operador.clear()

        if tipo == 'numero':
            for label_op, valor in FiltroBusca.OPERADORES_NUM.items():
                self._combo_operador.addItem(label_op, valor)
        else:
            for label_op, valor in FiltroBusca.OPERADORES_TEXT.items():
                self._combo_operador.addItem(label_op, valor)

    def _on_adicionar_filtro(self) -> None:
        """Adiciona um novo filtro."""
        campo = self._combo_campo.currentData()
        operador = self._combo_operador.currentData()
        valor = self._valor_input.text()

        if not valor:
            return

        filtro = FiltroBusca(
            campo=campo,
            operador=operador,
            valor=valor
        )

        self._filtros_ativos.append(filtro)

        # Cria widget visual
        widget = FiltroItemWidget(filtro)
        widget.removido.connect(lambda: self._remover_filtro(widget, filtro))

        self._filtros_layout.insertWidget(
            self._filtros_layout.count() - 1,
            widget
        )
        self._filtros_widgets.append(widget)

        self._valor_input.clear()

        self.filtro_adicionado.emit(filtro)

    def _remover_filtro(self, widget: FiltroItemWidget, filtro: FiltroBusca) -> None:
        """Remove um filtro."""
        if filtro in self._filtros_ativos:
            indice = self._filtros_ativos.index(filtro)
            self._filtros_ativos.remove(filtro)
            self._filtros_widgets.remove(widget)
            widget.deleteLater()

            self.filtro_removido.emit(indice)

    def _on_buscar(self) -> None:
        """Executa a busca."""
        texto = self._busca_input.text()
        self.busca_solicitada.emit(texto, self._filtros_ativos.copy())

    def _on_limpar(self) -> None:
        """Limpa busca e filtros."""
        self._busca_input.clear()

        for widget in self._filtros_widgets:
            widget.deleteLater()

        self._filtros_ativos.clear()
        self._filtros_widgets.clear()

        self.limpar_solicitado.emit()

    def _mostrar_menu_presets(self) -> None:
        """Mostra menu de presets."""
        menu = QMenu(self)

        action_carregar = menu.addAction("📂 Carregar Preset")
        action_carregar.triggered.connect(self._carregar_preset)

        action_gerenciar = menu.addAction("⚙️ Gerenciar Presets")
        action_gerenciar.triggered.connect(self._gerenciar_presets)

        menu.exec(self._btn_presets.mapToGlobal(self._btn_presets.rect().bottomLeft()))

    def _carregar_preset(self) -> None:
        """Abre diálogo para carregar preset."""
        # Emite sinal para que a janela principal trate
        self.preset_carregado.emit("")

    def _salvar_preset(self) -> None:
        """Salva os filtros atuais como preset."""
        if not self._filtros_ativos:
            QMessageBox.warning(self, "Aviso", "Não há filtros ativos para salvar.")
            return

        nome, ok = QInputDialog.getText(
            self, "Salvar Preset", "Nome do preset:"
        )

        if ok and nome:
            descricao, ok2 = QInputDialog.getText(
                self, "Descrição", "Descrição (opcional):"
            )

            if ok2:
                self.preset_salvo.emit(nome, descricao)

    def _gerenciar_presets(self) -> None:
        """Abre diálogo de gerenciamento de presets."""
        self.preset_gerenciar.emit()

    # Sinais adicionais
    preset_salvo = Signal(str, str)
    preset_gerenciar = Signal()

    @property
    def filtros_ativos(self) -> List[FiltroBusca]:
        """Retorna os filtros ativos."""
        return self._filtros_ativos.copy()

    @property
    def texto_busca(self) -> str:
        """Retorna o texto de busca."""
        return self._busca_input.text()

    def definir_texto_busca(self, texto: str) -> None:
        """Define o texto de busca."""
        self._busca_input.setText(texto)
