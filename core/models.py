"""Modelos de dados para suportes AutoCAD."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class SuporteData:
    """
    Representa um suporte de tubulação no AutoCAD.

    Attributes:
        tag: Identificador da posição (ex: "POS-001")
        tipo: Tipo do suporte (ex: "SP_EP-01-A")
        posicao_x: Coordenada X
        posicao_y: Coordenada Y
        posicao_z: Coordenada Z
        handle: Handle único do objeto no AutoCAD
        propriedades: Dicionário de propriedades dinâmicas do bloco
        layer: Camada do objeto
        selecionado: Indica se o suporte está selecionado para edição em lote
    """
    tag: str
    tipo: str
    posicao_x: float
    posicao_y: float
    posicao_z: float
    handle: str
    propriedades: Dict[str, Any] = field(default_factory=dict)
    layer: str = ""
    selecionado: bool = False

    def __post_init__(self):
        """Normaliza dados após inicialização."""
        self.tag = str(self.tag).strip()
        self.tipo = str(self.tipo).strip()
        self.handle = str(self.handle).strip()

    @property
    def posicao(self) -> str:
        """Retorna a posição como string formatada."""
        return f"({self.posicao_x:.2f}, {self.posicao_y:.2f}, {self.posicao_z:.2f})"

    def obter_propriedade(self, nome: str) -> Optional[Any]:
        """
        Obtém uma propriedade dinâmica pelo nome.

        Args:
            nome: Nome da propriedade

        Returns:
            Valor da propriedade ou None se não existir
        """
        return self.propriedades.get(nome)

    def definir_propriedade(self, nome: str, valor: Any) -> None:
        """
        Define uma propriedade dinâmica.

        Args:
            nome: Nome da propriedade
            valor: Valor a ser definido
        """
        self.propriedades[nome] = valor

    def listar_nomes_propriedades(self) -> List[str]:
        """
        Retorna lista de nomes de propriedades dinâmicas.

        Returns:
            Lista ordenada de nomes de propriedades
        """
        # Excluir 'Origin' que não é uma propriedade editável
        return sorted([p for p in self.propriedades.keys() if p != "Origin"])

    def to_dict(self) -> Dict[str, Any]:
        """
        Converte para dicionário.

        Returns:
            Dicionário com todos os dados do suporte
        """
        return {
            'tag': self.tag,
            'tipo': self.tipo,
            'posicao_x': self.posicao_x,
            'posicao_y': self.posicao_y,
            'posicao_z': self.posicao_z,
            'handle': self.handle,
            'layer': self.layer,
            'propriedades': self.propriedades.copy(),
            'selecionado': self.selecionado
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SuporteData':
        """
        Cria instância a partir de dicionário.

        Args:
            data: Dicionário com dados do suporte

        Returns:
            Nova instância de SuporteData
        """
        return cls(
            tag=data.get('tag', ''),
            tipo=data.get('tipo', ''),
            posicao_x=float(data.get('posicao_x', 0)),
            posicao_y=float(data.get('posicao_y', 0)),
            posicao_z=float(data.get('posicao_z', 0)),
            handle=data.get('handle', ''),
            propriedades=data.get('propriedades', {}),
            layer=data.get('layer', ''),
            selecionado=data.get('selecionado', False)
        )


@dataclass
class FiltroBusca:
    """
    Representa um filtro de busca.

    Attributes:
        campo: Campo a filtrar (tag, tipo, propriedade, etc.)
        operador: Operador (contem, inicia_termina, igual, maior, menor, entre)
        valor: Valor do filtro
        valor_secundario: Valor secundário para operador 'entre'
    """
    campo: str
    operador: str
    valor: Any
    valor_secundario: Optional[Any] = None

    OPERADORES_TEXT = {
        'contem': 'Contém',
        'nao_contem': 'Não Contém',
        'inicia_com': 'Inicia Com',
        'termina_com': 'Termina Com',
        'igual': 'Igual',
        'diferente': 'Diferente',
    }

    OPERADORES_NUM = {
        'igual': 'Igual',
        'maior': 'Maior Que',
        'menor': 'Menor Que',
        'maior_igual': 'Maior ou Igual',
        'menor_igual': 'Menor ou Igual',
        'entre': 'Entre',
    }

    def verificar(self, suporte: SuporteData) -> bool:
        """
        Verifica se o suporte atende ao filtro.

        Args:
            suporte: Suporte a verificar

        Returns:
            True se atende ao filtro
        """
        valor_alvo = None

        # Determina o valor alvo baseado no campo
        if self.campo == 'tag':
            valor_alvo = suporte.tag
        elif self.campo == 'tipo':
            valor_alvo = suporte.tipo
        elif self.campo == 'layer':
            valor_alvo = suporte.layer
        elif self.campo in suporte.propriedades:
            valor_alvo = suporte.propriedades[self.campo]
        else:
            return False

        # Aplica operação baseada no tipo de valor
        if isinstance(valor_alvo, (int, float)) and self.operador in self.OPERADORES_NUM.values():
            return self._verificar_numerico(valor_alvo)
        else:
            return self._verificar_texto(str(valor_alvo))

    def _verificar_texto(self, valor_alvo: str) -> bool:
        """Verifica filtro para valores de texto."""
        valor_filtro = str(self.valor).lower()
        valor_alvo_lower = valor_alvo.lower()

        if self.operador == 'contem' or self.operador == self.OPERADORES_TEXT['contem']:
            return valor_filtro in valor_alvo_lower
        elif self.operador == 'nao_contem' or self.operador == self.OPERADORES_TEXT['nao_contem']:
            return valor_filtro not in valor_alvo_lower
        elif self.operador == 'inicia_com' or self.operador == self.OPERADORES_TEXT['inicia_com']:
            return valor_alvo_lower.startswith(valor_filtro)
        elif self.operador == 'termina_com' or self.operador == self.OPERADORES_TEXT['termina_com']:
            return valor_alvo_lower.endswith(valor_filtro)
        elif self.operador == 'igual' or self.operador == self.OPERADORES_TEXT['igual']:
            return valor_alvo_lower == valor_filtro
        elif self.operador == 'diferente' or self.operador == self.OPERADORES_TEXT['diferente']:
            return valor_alvo_lower != valor_filtro
        return False

    def _verificar_numerico(self, valor_alvo: float) -> bool:
        """Verifica filtro para valores numéricos."""
        try:
            valor_filtro = float(self.valor)
        except (ValueError, TypeError):
            return False

        if self.operador == 'igual' or self.operador == self.OPERADORES_NUM['igual']:
            return valor_alvo == valor_filtro
        elif self.operador == 'maior' or self.operador == self.OPERADORES_NUM['maior']:
            return valor_alvo > valor_filtro
        elif self.operador == 'menor' or self.operador == self.OPERADORES_NUM['menor']:
            return valor_alvo < valor_filtro
        elif self.operador == 'maior_igual' or self.operador == self.OPERADORES_NUM['maior_igual']:
            return valor_alvo >= valor_filtro
        elif self.operador == 'menor_igual' or self.operador == self.OPERADORES_NUM['menor_igual']:
            return valor_alvo <= valor_filtro
        elif self.operador == 'entre' or self.operador == self.OPERADORES_NUM['entre']:
            if self.valor_secundario is None:
                return False
            try:
                valor_sec = float(self.valor_secundario)
                return min(valor_filtro, valor_sec) <= valor_alvo <= max(valor_filtro, valor_sec)
            except (ValueError, TypeError):
                return False
        return False

    def __str__(self) -> str:
        """Representação textual do filtro."""
        operador_label = self.OPERADORES_TEXT.get(self.operador, self.OPERADORES_NUM.get(self.operador, self.operador))
        if self.operador == 'entre' and self.valor_secundario:
            return f"{self.campo} {operador_label} {self.valor} e {self.valor_secundario}"
        return f"{self.campo} {operador_label} {self.valor}"
