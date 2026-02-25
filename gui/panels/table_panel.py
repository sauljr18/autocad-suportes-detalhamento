"""Painel de tabela de suportes."""

from typing import List, Optional

from PySide6.QtCore import Qt, Signal, QModelIndex
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QPushButton, QTableView,
    QVBoxLayout, QWidget, QHeaderView, QAbstractItemView,
    QMenu, QDialog, QDialogButtonBox, QTextEdit
)

from gui.models.suporte_table_model import SuporteTableModel
from core.models import SuporteData


class TablePanel(QWidget):
    """
    Painel com tabela de suportes.

    Sinais:
        suporte_selecionado: Emitido quando suporte é selecionado
        zoom_solicitado: Emitido para zoom no AutoCAD
        editar_solicitado: Emitido para editar suporte
        selecao_mudou: Emitido quando seleção muda (qtd)
        atualizar_solicitado: Emitido para recarregar dados
    """

    suporte_selecionado = Signal(object)
    zoom_solicitado = Signal(str)
    editar_solicitado = Signal(str)
    selecao_mudou = Signal(int)
    atualizar_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self._model = SuporteTableModel(self)
        self._setup_ui()

    def _setup_ui(self) -> None:
        """Configura a UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Barra de ferramentas
        toolbar_layout = QHBoxLayout()

        self._label_contagem = QLabel("0 suportes")
        self._label_selecionados = QLabel("0 selecionados")

        btn_atualizar = QPushButton("🔄 Atualizar")
        btn_atualizar.clicked.connect(self.atualizar_solicitado.emit)

        btn_selecionar_todos = QPushButton("✓ Todos")
        btn_selecionar_todos.clicked.connect(self._selecionar_todos)

        btn_limpar_selecao = QPushButton("✗ Limpar")
        btn_limpar_selecao.clicked.connect(self._limpar_selecao)

        toolbar_layout.addWidget(self._label_contagem)
        toolbar_layout.addWidget(self._label_selecionados)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(btn_atualizar)
        toolbar_layout.addWidget(btn_selecionar_todos)
        toolbar_layout.addWidget(btn_limpar_selecao)

        layout.addLayout(toolbar_layout)

        # Tabela
        self._table = QTableView()
        self._table.setModel(self._model)

        # Configurações da tabela
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setSortingEnabled(False)

        # Ajuste de colunas
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Checkbox
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # TAG
        header.setSectionResizeMode(2, QHeaderView.Interactive)       # Tipo
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # X
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Y
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Z
        header.setSectionResizeMode(6, QHeaderView.Interactive)       # Camada
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Ações

        # Altura das linhas
        self._table.verticalHeader().setDefaultSectionSize(24)

        # Menu de contexto
        self._table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._mostrar_menu_contexto)

        # Conecta sinais
        self._table.clicked.connect(self._on_clicked)
        self._table.doubleClicked.connect(self._on_double_clicked)
        selection_model = self._table.selectionModel()
        if selection_model:
            selection_model.selectionChanged.connect(self._on_selection_changed)

        layout.addWidget(self._table)

    def _on_clicked(self, index: QModelIndex) -> None:
        """Trata clique na tabela."""
        if index.column() == 0:  # Checkbox
            self.model().setData(index, index.data(Qt.CheckStateRole) == Qt.Unchecked, Qt.CheckStateRole)
            self._atualizar_label_selecao()

    def _on_double_clicked(self, index: QModelIndex) -> None:
        """Trata duplo clique para editar."""
        suporte = self._model.obter_suporte(index.row())
        if suporte:
            self.editar_solicitado.emit(suporte.handle)

    def _on_selection_changed(self) -> None:
        """Trata mudança de seleção."""
        indexes = self._table.selectionModel().selectedRows()
        if indexes:
            row = indexes[0].row()
            suporte = self._model.obter_suporte(row)
            if suporte:
                self.suporte_selecionado.emit(suporte)

    def _atualizar_label_selecao(self) -> None:
        """Atualiza label de selecionados."""
        count = self._model.contar_selecionados()
        self._label_selecionados.setText(f"{count} selecionados")
        self.selecao_mudou.emit(count)

    def _selecionar_todos(self) -> None:
        """Seleciona todos."""
        self._model.selecionar_todos(True)
        self._atualizar_label_selecao()

    def _limpar_selecao(self) -> None:
        """Limpa seleção."""
        self._model.limpar_selecao()
        self._atualizar_label_selecao()

    def _mostrar_menu_contexto(self, pos) -> None:
        """Mostra menu de contexto."""
        index = self._table.indexAt(pos)
        if not index.isValid():
            return

        suporte = self._model.obter_suporte(index.row())
        if not suporte:
            return

        menu = QMenu(self)

        action_zoom = menu.addAction("🔍 Zoom no AutoCAD")
        action_zoom.triggered.connect(lambda: self.zoom_solicitado.emit(suporte.handle))

        action_editar = menu.addAction("✏️ Editar Propriedades")
        action_editar.triggered.connect(lambda: self.editar_solicitado.emit(suporte.handle))

        menu.addSeparator()

        # Adiciona menu com propriedades
        props_menu = menu.addMenu("📋 Propriedades")
        for nome in sorted(suporte.listar_nomes_propriedades()):
            valor = suporte.obter_propriedade(nome)
            action = props_menu.addAction(f"{nome}: {valor}")
            action.setEnabled(False)

        menu.exec(self._table.mapToGlobal(pos))

    def model(self) -> SuporteTableModel:
        """Retorna o modelo da tabela."""
        return self._model

    def atualizar_dados(self, suportes: List[SuporteData]) -> None:
        """
        Atualiza os dados da tabela.

        Args:
            suportes: Nova lista de suportes
        """
        self._model.atualizar_dados(suportes)
        self._label_contagem.setText(f"{len(suportes)} suportes")
        self._atualizar_label_selecao()

    def adicionar_suporte(self, suporte: SuporteData) -> None:
        """Adiciona um suporte à tabela."""
        self._model.adicionar_suporte(suporte)
        self._label_contagem.setText(f"{self._model.rowCount()} suportes")

    def limpar(self) -> None:
        """Limpa a tabela."""
        self._model.limpar()
        self._label_contagem.setText("0 suportes")
        self._label_selecionados.setText("0 selecionados")

    def obter_suporte_selecionado(self) -> Optional[SuporteData]:
        """Retorna o suporte atualmente selecionado."""
        indexes = self._table.selectionModel().selectedRows()
        if indexes:
            row = indexes[0].row()
            return self._model.obter_suporte(row)
        return None

    def obter_selecionados(self) -> List[SuporteData]:
        """Retorna todos os suportes com checkbox marcado."""
        return self._model.obter_selecionados()

    def selecionar_por_handle(self, handle: str) -> bool:
        """
        Seleciona um suporte pelo handle.

        Returns:
            True se encontrado e selecionado
        """
        for row in range(self._model.rowCount()):
            suporte = self._model.obter_suporte(row)
            if suporte and suporte.handle == handle:
                self._table.selectRow(row)
                return True
        return False

    def ordenar_por_tag(self) -> None:
        """Ordena a tabela por TAG."""
        self._model.ordenar_por_tag()

    def ordenar_por_tipo(self) -> None:
        """Ordena a tabela por tipo."""
        self._model.ordenar_por_tipo()
