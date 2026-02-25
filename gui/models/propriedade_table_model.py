"""Modelo de tabela para propriedades de suporte."""

from typing import Any, Dict, List, Optional

from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor


class Propriedade:
    """Representa uma propriedade editável."""

    def __init__(
        self,
        nome: str,
        valor: Any,
        minimo: Optional[float] = None,
        maximo: Optional[float] = None,
        readonly: bool = False
    ):
        self.nome = nome
        self.valor = valor
        self.minimo = minimo
        self.maximo = maximo
        self.readonly = readonly


class PropriedadeTableModel(QAbstractTableModel):
    """
    Model para QTableView de propriedades.

    Colunas:
    - Propriedade
    - Valor Atual
    - Limites (min/max)
    """

    COL_NOME = 0
    COL_VALOR = 1
    COL_LIMITES = 2

    COLUMN_COUNT = 3

    def __init__(self, parent=None):
        """Inicializa o modelo."""
        super().__init__(parent)
        self._propriedades: List[Propriedade] = []
        self._headers = ["Propriedade", "Valor", "Limites"]

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Retorna o número de linhas."""
        return len(self._propriedades)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        """Retorna o número de colunas."""
        return self.COLUMN_COUNT

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        """Retorna dados para uma célula."""
        if not index.isValid() or index.row() >= len(self._propriedades):
            return None

        prop = self._propriedades[index.row()]
        col = index.column()

        if role == Qt.DisplayRole or role == Qt.EditRole:
            if col == self.COL_NOME:
                return prop.nome
            elif col == self.COL_VALOR:
                return str(prop.valor) if prop.valor is not None else ""
            elif col == self.COL_LIMITES:
                if prop.minimo is not None and prop.maximo is not None:
                    return f"{prop.minimo} - {prop.maximo}"
                elif prop.minimo is not None:
                    return f">= {prop.minimo}"
                elif prop.maximo is not None:
                    return f"<= {prop.maximo}"
                else:
                    return ""

        elif role == Qt.TextAlignmentRole:
            if col == self.COL_VALOR or col == self.COL_LIMITES:
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.BackgroundRole:
            # Linhas alternadas
            if index.row() % 2 == 0:
                return QColor(240, 240, 240)
            # Propriedades readonly em cinza claro
            if prop.readonly:
                return QColor(230, 230, 230)

        elif role == Qt.ForegroundRole:
            if prop.readonly:
                return QColor(128, 128, 128)

        elif role == Qt.ToolTipRole:
            if prop.readonly:
                return "Propriedade somente leitura"
            if col == self.COL_LIMITES and prop.minimo is not None and prop.maximo is not None:
                return f"Valor mínimo: {prop.minimo}\nValor máximo: {prop.maximo}"

        elif role == Qt.UserRole:
            return prop

        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        """Define dados para uma célula."""
        if not index.isValid() or index.row() >= len(self._propriedades):
            return False

        prop = self._propriedades[index.row()]

        if role == Qt.EditRole and index.column() == self.COL_VALOR:
            if prop.readonly:
                return False

            # Tenta converter para o tipo apropriado
            try:
                if isinstance(prop.valor, float):
                    novo_valor = float(value)
                elif isinstance(prop.valor, int):
                    novo_valor = int(value)
                else:
                    novo_valor = str(value)

                # Verifica limites
                if prop.minimo is not None and novo_valor < prop.minimo:
                    return False
                if prop.maximo is not None and novo_valor > prop.maximo:
                    return False

                prop.valor = novo_valor
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                return True

            except (ValueError, TypeError):
                return False

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

        # Coluna de valor é editável se não for readonly
        if index.column() == self.COL_VALOR:
            prop = self._propriedades[index.row()]
            if not prop.readonly:
                flags |= Qt.ItemIsEditable

        return flags

    def atualizar_dados(self, propriedades: Dict[str, Any]) -> None:
        """
        Atualiza os dados do modelo.

        Args:
            propriedades: Dicionário de propriedades no formato:
                {nome: {'valor': ..., 'min': ..., 'max': ..., 'readonly': ...}}
        """
        self.beginResetModel()
        self._propriedades.clear()

        for nome, dados in propriedades.items():
            if isinstance(dados, dict):
                prop = Propriedade(
                    nome=nome,
                    valor=dados.get('valor'),
                    minimo=dados.get('min'),
                    maximo=dados.get('max'),
                    readonly=dados.get('readonly', False)
                )
            else:
                prop = Propriedade(nome=nome, valor=dados)

            self._propriedades.append(prop)

        # Ordena por nome
        self._propriedades.sort(key=lambda p: p.nome)

        self.endResetModel()

    def atualizar_lista(self, propriedades: List[Propriedade]) -> None:
        """
        Atualiza com uma lista de Propriedade.

        Args:
            propriedades: Lista de Propriedade
        """
        self.beginResetModel()
        self._propriedades = propriedades
        self.endResetModel()

    def limpar(self) -> None:
        """Limpa todos os dados do modelo."""
        self.beginResetModel()
        self._propriedades.clear()
        self.endResetModel()

    def obter_propriedade(self, row: int) -> Optional[Propriedade]:
        """
        Obtém uma propriedade por linha.

        Args:
            row: Índice da linha

        Returns:
            Propriedade ou None
        """
        if 0 <= row < len(self._propriedades):
            return self._propriedades[row]
        return None

    def obter_valor(self, row: int) -> Optional[Any]:
        """
        Obtém o valor de uma propriedade.

        Args:
            row: Índice da linha

        Returns:
            Valor ou None
        """
        prop = self.obter_propriedade(row)
        return prop.valor if prop else None

    def definir_valor(self, row: int, valor: Any) -> bool:
        """
        Define o valor de uma propriedade.

        Args:
            row: Índice da linha
            valor: Novo valor

        Returns:
            True se definido com sucesso
        """
        if 0 <= row < len(self._propriedades):
            prop = self._propriedades[row]

            if prop.readonly:
                return False

            # Verifica limites
            if prop.minimo is not None and valor < prop.minimo:
                return False
            if prop.maximo is not None and valor > prop.maximo:
                return False

            prop.valor = valor

            index = self.createIndex(row, self.COL_VALOR)
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])

            return True

        return False

    def para_dicionario(self) -> Dict[str, Any]:
        """
        Converte para dicionário.

        Returns:
            Dicionário {nome: valor}
        """
        return {prop.nome: prop.valor for prop in self._propriedades}

    def contem_propriedade(self, nome: str) -> bool:
        """
        Verifica se contém uma propriedade.

        Args:
            nome: Nome da propriedade

        Returns:
            True se contém
        """
        return any(p.nome == nome for p in self._propriedades)

    def obter_indice_por_nome(self, nome: str) -> int:
        """
        Obtém o índice de uma propriedade pelo nome.

        Args:
            nome: Nome da propriedade

        Returns:
            Índice ou -1 se não encontrado
        """
        for i, prop in enumerate(self._propriedades):
            if prop.nome == nome:
                return i
        return -1
