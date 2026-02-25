"""Modelo de tabela para suportes."""

from typing import Any, List, Optional

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor

from core.models import SuporteData


class SuporteTableModel(QAbstractTableModel):
    """
    Model para QTableView de suportes.

    Colunas:
    - Seleção (checkbox)
    - TAG
    - Tipo
    - Posição X
    - Posição Y
    - Posição Z
    - Camada
    - Ações
    """

    COL_CHECKBOX = 0
    COL_TAG = 1
    COL_TIPO = 2
    COL_X = 3
    COL_Y = 4
    COL_Z = 5
    COL_CAMADA = 6
    COL_ACOES = 7

    COLUMN_COUNT = 8

    def __init__(self, parent=None):
        """Inicializa o modelo."""
        super().__init__(parent)
        self._suportes: List[SuporteData] = []
        self._headers = [
            "✓", "TAG", "Tipo", "X", "Y", "Z", "Camada", "Ações"
        ]

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Retorna o número de linhas."""
        return len(self._suportes)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Retorna o número de colunas."""
        return self.COLUMN_COUNT

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        """Retorna dados para uma célula."""
        if not index.isValid() or index.row() >= len(self._suportes):
            return None

        suporte = self._suportes[index.row()]
        col = index.column()

        if role == Qt.DisplayRole or role == Qt.EditRole:
            if col == self.COL_TAG:
                return suporte.tag
            elif col == self.COL_TIPO:
                return suporte.tipo
            elif col == self.COL_X:
                return f"{suporte.posicao_x:.2f}"
            elif col == self.COL_Y:
                return f"{suporte.posicao_y:.2f}"
            elif col == self.COL_Z:
                return f"{suporte.posicao_z:.2f}"
            elif col == self.COL_CAMADA:
                return suporte.layer

        elif role == Qt.CheckStateRole and col == self.COL_CHECKBOX:
            return Qt.Checked if suporte.selecionado else Qt.Unchecked

        elif role == Qt.TextAlignmentRole:
            if col in (self.COL_X, self.COL_Y, self.COL_Z):
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.BackgroundRole:
            # Linhas alternadas
            if index.row() % 2 == 0:
                return QColor(240, 240, 240)

        elif role == Qt.ToolTipRole:
            if col == self.COL_TAG:
                return f"Handle: {suporte.handle}"
            elif col == self.COL_TIPO:
                return f"Propriedades: {len(suporte.propriedades)}"

        elif role == Qt.UserRole:
            # Retorna o objeto completo
            return suporte

        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        """Define dados para uma célula."""
        if not index.isValid() or index.row() >= len(self._suportes):
            return False

        suporte = self._suportes[index.row()]

        if role == Qt.CheckStateRole and index.column() == self.COL_CHECKBOX:
            suporte.selecionado = (value == Qt.Checked)
            self.dataChanged.emit(index, index, [Qt.CheckStateRole])
            return True

        return False

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.DisplayRole
    ) -> Any:
        """Retorna dados do cabeçalho."""
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            if section < len(self._headers):
                return self._headers[section]
        return None

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        """Retorna flags para um item."""
        if not index.isValid():
            return Qt.NoItemFlags

        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable

        # Coluna de checkbox é editável
        if index.column() == self.COL_CHECKBOX:
            flags |= Qt.ItemIsUserCheckable

        return flags

    def atualizar_dados(self, suportes: List[SuporteData]) -> None:
        """
        Atualiza os dados do modelo.

        Args:
            suportes: Nova lista de suportes
        """
        self.beginResetModel()
        self._suportes = suportes
        self.endResetModel()

    def adicionar_suporte(self, suporte: SuporteData) -> None:
        """
        Adiciona um suporte ao modelo.

        Args:
            suporte: Suporte a adicionar
        """
        row = len(self._suportes)
        self.beginInsertRows(QModelIndex(), row, row)
        self._suportes.append(suporte)
        self.endInsertRows()

    def remover_suporte(self, row: int) -> None:
        """
        Remove um suporte do modelo.

        Args:
            row: Índice da linha
        """
        if 0 <= row < len(self._suportes):
            self.beginRemoveRows(QModelIndex(), row, row)
            del self._suportes[row]
            self.endRemoveRows()

    def limpar(self) -> None:
        """Limpa todos os dados do modelo."""
        self.beginResetModel()
        self._suportes.clear()
        self.endResetModel()

    def obter_suporte(self, row: int) -> Optional[SuporteData]:
        """
        Obtém um suporte por linha.

        Args:
            row: Índice da linha

        Returns:
            SuporteData ou None
        """
        if 0 <= row < len(self._suportes):
            return self._suportes[row]
        return None

    def obter_suporte_por_handle(self, handle: str) -> Optional[SuporteData]:
        """
        Obtém um suporte pelo handle.

        Args:
            handle: Handle do suporte

        Returns:
            SuporteData ou None
        """
        for suporte in self._suportes:
            if suporte.handle == handle:
                return suporte
        return None

    def obter_selecionados(self) -> List[SuporteData]:
        """
        Obtém todos os suportes selecionados.

        Returns:
            Lista de SuporteData selecionados
        """
        return [s for s in self._suportes if s.selecionado]

    def selecionar_todos(self, selecionado: bool = True) -> None:
        """
        Seleciona ou desseleciona todos os suportes.

        Args:
            selecionado: True para selecionar, False para desselecionar
        """
        for suporte in self._suportes:
            suporte.selecionado = selecionado

        # Emite sinal de mudança para toda a coluna de checkbox
        self.dataChanged.emit(
            self.index(0, self.COL_CHECKBOX),
            self.index(len(self._suportes) - 1, self.COL_CHECKBOX),
            [Qt.CheckStateRole]
        )

    def inverter_selecao(self) -> None:
        """Inverte a seleção de todos os suportes."""
        for suporte in self._suportes:
            suporte.selecionado = not suporte.selecionado

        self.dataChanged.emit(
            self.index(0, self.COL_CHECKBOX),
            self.index(len(self._suportes) - 1, self.COL_CHECKBOX),
            [Qt.CheckStateRole]
        )

    def limpar_selecao(self) -> None:
        """Limpa a seleção de todos os suportes."""
        self.selecionar_todos(False)

    def contar_selecionados(self) -> int:
        """
        Conta quantos suportes estão selecionados.

        Returns:
            Número de suportes selecionados
        """
        return sum(1 for s in self._suportes if s.selecionado)

    def ordenar_por_tag(self) -> None:
        """Ordena os suportes por tag."""
        self.beginResetModel()
        self._suportes.sort(key=lambda s: s.tag)
        self.endResetModel()

    def ordenar_por_tipo(self) -> None:
        """Ordena os suportes por tipo."""
        self.beginResetModel()
        self._suportes.sort(key=lambda s: s.tipo)
        self.endResetModel()
