"""Painel de edição de propriedades."""

from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDoubleSpinBox, QFrame, QGroupBox,
    QHBoxLayout, QLabel, QLineEdit, QPushButton, QSpinBox,
    QTableView, QVBoxLayout, QWidget, QHeaderView, QMessageBox
)

from gui.models.propriedade_table_model import PropriedadeTableModel
from core.models import SuporteData


class EditPanel(QWidget):
    """
    Painel de edição de propriedades do suporte.

    Sinais:
        valor_alterado: Emitido quando valor é alterado (handle, propriedade, valor)
        aplicar_lote: Emitido para aplicar valor em lote
    """

    valor_alterado = Signal(str, str, object)
    aplicar_lote = Signal(str, object)

    def __init__(self, parent=None):
        super().__init__(parent)

        self._suporte_atual: Optional[SuporteData] = None
        self._model = PropriedadeTableModel(self)
        self._modo_lote: bool = False

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Configura a UI."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Título
        self._label_titulo = QLabel("Selecione um suporte para editar")
        self._label_titulo.setStyleSheet("font-weight: bold; font-size: 12px;")
        layout.addWidget(self._label_titulo)

        # Grupo de informações
        info_group = QGroupBox("Informações")
        info_layout = QVBoxLayout()

        self._label_tag = QLabel("TAG: -")
        self._label_tipo = QLabel("Tipo: -")
        self._label_posicao = QLabel("Posição: -")

        info_layout.addWidget(self._label_tag)
        info_layout.addWidget(self._label_tipo)
        info_layout.addWidget(self._label_posicao)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # Tabela de propriedades
        props_group = QGroupBox("Propriedades Dinâmicas")
        props_layout = QVBoxLayout()

        self._table = QTableView()
        self._table.setModel(self._model)
        self._table.setSelectionBehavior(QTableView.SelectRows)
        self._table.setSelectionMode(QTableView.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)

        # Ajuste de colunas
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        props_layout.addWidget(self._table)
        props_group.setLayout(props_layout)
        layout.addWidget(props_group)

        # Controles de edição
        edit_layout = QHBoxLayout()

        # Combo para seleção de propriedade (modo lote)
        self._combo_propriedade = QComboBox()
        self._combo_propriedade.setEnabled(False)
        self._combo_propriedade.setMinimumWidth(150)

        # Campo de valor
        self._valor_input = QLineEdit()
        self._valor_input.setPlaceholderText("Novo valor...")
        self._valor_input.setEnabled(False)

        # Botão aplicar
        self._btn_aplicar = QPushButton("Aplicar")
        self._btn_aplicar.setEnabled(False)
        self._btn_aplicar.clicked.connect(self._on_aplicar)

        edit_layout.addWidget(QLabel("Propriedade:"))
        edit_layout.addWidget(self._combo_propriedade)
        edit_layout.addWidget(QLabel("Valor:"))
        edit_layout.addWidget(self._valor_input)
        edit_layout.addWidget(self._btn_aplicar)

        layout.addLayout(edit_layout)

        # Checkbox modo lote
        self._check_modo_lote = QCheckBox("Editar em Lote (aplicar a todos selecionados)")
        self._check_modo_lote.toggled.connect(self._on_modo_lote_changed)
        layout.addWidget(self._check_modo_lote)

        layout.addStretch()

    def definir_suporte(self, suporte: Optional[SuporteData]) -> None:
        """
        Define o suporte a ser editado.

        Args:
            suporte: SuporteData ou None para limpar
        """
        self._suporte_atual = suporte

        if suporte is None:
            self._label_titulo.setText("Selecione um suporte para editar")
            self._label_tag.setText("TAG: -")
            self._label_tipo.setText("Tipo: -")
            self._label_posicao.setText("Posição: -")
            self._model.limpar()
            self._desabilitar_edicao()
            return

        # Atualiza informações
        self._label_titulo.setText(f"Editando: {suporte.tag}")
        self._label_tag.setText(f"TAG: {suporte.tag}")
        self._label_tipo.setText(f"Tipo: {suporte.tipo}")
        self._label_posicao.setText(f"Posição: {suporte.posicao}")

        # Carrega propriedades
        props = {}
        for nome in suporte.listar_nomes_propriedades():
            valor = suporte.obter_propriedade(nome)
            props[nome] = {'valor': valor}

        self._model.atualizar_dados(props)

        # Atualiza combo
        self._combo_propriedade.clear()
        for nome in suporte.listar_nomes_propriedades():
            self._combo_propriedade.addItem(nome)

        # Habilita edição
        self._habilitar_edicao()

    def definir_propriedades(self, propriedades: Dict[str, Any]) -> None:
        """
        Define as propriedades diretamente (para edição).

        Args:
            propriedades: Dicionário de propriedades
        """
        self._model.atualizar_dados(propriedades)

    def definir_lista_propriedades(self, lista: List[str]) -> None:
        """
        Define a lista de propriedades disponíveis (modo lote).

        Args:
            lista: Lista de nomes de propriedades
        """
        self._combo_propriedade.clear()
        self._combo_propriedade.addItems(sorted(lista))

    def _habilitar_edicao(self) -> None:
        """Habilita controles de edição."""
        if self._modo_lote:
            self._combo_propriedade.setEnabled(True)
            self._valor_input.setEnabled(True)
            self._btn_aplicar.setEnabled(True)
        else:
            self._combo_propriedade.setEnabled(False)
            self._valor_input.setEnabled(False)
            self._btn_aplicar.setEnabled(False)

        self._table.setEnabled(True)

    def _desabilitar_edicao(self) -> None:
        """Desabilita controles de edição."""
        self._combo_propriedade.setEnabled(False)
        self._valor_input.setEnabled(False)
        self._btn_aplicar.setEnabled(False)
        self._table.setEnabled(False)

    def _on_modo_lote_changed(self, checked: bool) -> None:
        """Trata mudança do modo lote."""
        self._modo_lote = checked

        if checked:
            self._label_titulo.setText("Modo de Edição em Lote")
            self._combo_propriedade.setEnabled(True)
            self._valor_input.setEnabled(True)
            self._btn_aplicar.setEnabled(True)
            self._table.setEnabled(False)
        else:
            if self._suporte_atual:
                self.definir_suporte(self._suporte_atual)

    def _on_aplicar(self) -> None:
        """Aplica a alteração."""
        if self._modo_lote:
            propriedade = self._combo_propriedade.currentText()
            valor = self._valor_input.text()

            if not propriedade or not valor:
                QMessageBox.warning(self, "Aviso", "Preencha a propriedade e o valor.")
                return

            self.aplicar_lote.emit(propriedade, valor)
        else:
            # Edição individual (da tabela)
            index = self._table.currentIndex()
            if index.isValid():
                propriedade = self._model.data(
                    self._model.index(index.row(), 0),
                    Qt.DisplayRole
                )
                valor = self._model.data(
                    self._model.index(index.row(), 1),
                    Qt.DisplayRole
                )

                if self._suporte_atual:
                    self.valor_alterado.emit(
                        self._suporte_atual.handle,
                        propriedade,
                        valor
                    )

    def obter_valor_alterado(self) -> Optional[tuple[str, object]]:
        """
        Obtém o valor a ser alterado dos controles.

        Returns:
            Tupla (propriedade, valor) ou None
        """
        propriedade = self._combo_propriedade.currentText()
        valor = self._valor_input.text()

        if not propriedade or not valor:
            return None

        # Tenta converter para número
        try:
            if '.' in valor:
                valor = float(valor)
            else:
                valor = int(valor)
        except ValueError:
            pass

        return (propriedade, valor)

    def model(self) -> PropriedadeTableModel:
        """Retorna o modelo da tabela."""
        return self._model

    @property
    def modo_lote(self) -> bool:
        """Retorna se está em modo lote."""
        return self._modo_lote

    def definir_modo_lote(self, modo: bool) -> None:
        """Define o modo lote."""
        self._check_modo_lote.setChecked(modo)

    def limpar(self) -> None:
        """Limpa o painel."""
        self.definir_suporte(None)
        self._check_modo_lote.setChecked(False)
