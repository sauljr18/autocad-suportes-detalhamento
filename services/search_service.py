"""Serviço de busca e filtros para suportes."""

import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from core.models import SuporteData, FiltroBusca
from core.repository import SuporteRepository
from .preset_manager import PresetManager
from .history_manager import HistoryManager


class SearchService:
    """
    Gerencia buscas e filtros de suportes.

    Coordena o repositório, gerenciador de presets e histórico
    para fornecer funcionalidade completa de busca.
    """

    def __init__(self, repository: SuporteRepository, data_dir: str = "data"):
        """
        Inicializa o serviço de busca.

        Args:
            repository: Repositório de suportes
            data_dir: Diretório para dados persistentes
        """
        self._repository = repository
        self._preset_manager = PresetManager(data_dir)
        self._history_manager = HistoryManager(data_dir)
        self._ultima_busca: List[SuporteData] = []
        self._filtros_ativos: List[FiltroBusca] = []

    @property
    def repository(self) -> SuporteRepository:
        """Retorna o repositório."""
        return self._repository

    @property
    def filtros_ativos(self) -> List[FiltroBusca]:
        """Retorna os filtros ativos."""
        return self._filtros_ativos.copy()

    @property
    def ultima_busca(self) -> List[SuporteData]:
        """Retorna o resultado da última busca."""
        return self._ultima_busca.copy()

    def buscar(
        self,
        texto_geral: str = "",
        filtros: Optional[List[FiltroBusca]] = None,
        salvar_historico: bool = True
    ) -> List[SuporteData]:
        """
        Executa uma busca com texto geral e/ou filtros.

        Args:
            texto_geral: Texto para busca em todas as colunas
            filtros: Lista de filtros específicos
            salvar_historico: Se True, salva no histórico

        Returns:
            Lista de suportes encontrados
        """
        if filtros is None:
            filtros = []

        self._filtros_ativos = filtros.copy()

        # Busca com filtros específicos
        if filtros:
            resultado = self._repository.buscar_por_filtro(filtros)
        else:
            resultado = self._repository.listar_todos()

        # Aplica filtro de texto geral se fornecido
        if texto_geral:
            resultado = self._filtrar_por_texto_geral(resultado, texto_geral)

        self._ultima_busca = resultado

        # Salva no histórico se solicitado
        if salvar_historico:
            self._salvar_busca_historico(texto_geral, filtros)

        return resultado

    def _filtrar_por_texto_geral(self, suportes: List[SuporteData], texto: str) -> List[SuporteData]:
        """
        Filtra suportes por texto geral (busca em todas as colunas).

        Args:
            suportes: Lista de suportes
            texto: Texto de busca

        Returns:
            Lista filtrada
        """
        texto_lower = texto.lower()
        resultado = []

        for suporte in suportes:
            # Busca em campos principais
            if (texto_lower in suporte.tag.lower() or
                texto_lower in suporte.tipo.lower() or
                texto_lower in suporte.layer.lower()):
                resultado.append(suporte)
                continue

            # Busca em propriedades dinâmicas
            encontrado = False
            for nome, valor in suporte.propriedades.items():
                if (texto_lower in nome.lower() or
                    texto_lower in str(valor).lower()):
                    encontrado = True
                    break

            if encontrado:
                resultado.append(suporte)

        return resultado

    def _salvar_busca_historico(self, texto_geral: str, filtros: List[FiltroBusca]) -> None:
        """Salva a busca no histórico."""
        dados_busca = {
            'texto_geral': texto_geral,
            'filtros': [self._filtro_to_dict(f) for f in filtros],
            'data_hora': datetime.now().isoformat(),
            'resultados': len(self._ultima_busca)
        }
        self._history_manager.adicionar(dados_busca)

    def _filtro_to_dict(self, filtro: FiltroBusca) -> Dict:
        """Converte filtro para dicionário."""
        return {
            'campo': filtro.campo,
            'operador': filtro.operador,
            'valor': filtro.valor,
            'valor_secundario': filtro.valor_secundario
        }

    def _dict_to_filtro(self, data: Dict) -> FiltroBusca:
        """Converte dicionário para filtro."""
        return FiltroBusca(
            campo=data['campo'],
            operador=data['operador'],
            valor=data['valor'],
            valor_secundario=data.get('valor_secundario')
        )

    def limpar_filtros(self) -> None:
        """Limpa os filtros ativos."""
        self._filtros_ativos.clear()
        self._ultima_busca = self._repository.listar_todos()

    # === Métodos de Preset ===

    def criar_preset(self, nome: str, descricao: str = "") -> Tuple[bool, str]:
        """
        Cria um preset com os filtros atuais.

        Args:
            nome: Nome do preset
            descricao: Descrição opcional

        Returns:
            Tupla (sucesso, mensagem)
        """
        if not self._filtros_ativos:
            return False, "Não há filtros ativos para salvar"

        dados_preset = {
            'nome': nome,
            'descricao': descricao,
            'filtros': [self._filtro_to_dict(f) for f in self._filtros_ativos],
            'data_criacao': datetime.now().isoformat()
        }

        return self._preset_manager.salvar(nome, dados_preset)

    def carregar_preset(self, nome: str) -> Tuple[bool, str, List[FiltroBusca]]:
        """
        Carrega um preset.

        Args:
            nome: Nome do preset

        Returns:
            Tupla (sucesso, mensagem, filtros)
        """
        sucesso, preset = self._preset_manager.carregar(nome)

        if not sucesso:
            return False, preset, []

        filtros = [self._dict_to_filtro(f) for f in preset.get('filtros', [])]
        self._filtros_ativos = filtros

        return True, f"Preset '{nome}' carregado", filtros

    def listar_presets(self) -> List[Dict[str, Any]]:
        """
        Lista todos os presets disponíveis.

        Returns:
            Lista de presets (nome, descricao, data_criacao)
        """
        return self._preset_manager.listar_todos()

    def deletar_preset(self, nome: str) -> Tuple[bool, str]:
        """
        Deleta um preset.

        Args:
            nome: Nome do preset

        Returns:
            Tupla (sucesso, mensagem)
        """
        return self._preset_manager.deletar(nome)

    # === Métodos de Histórico ===

    def obter_historico(self, limite: int = 50) -> List[Dict[str, Any]]:
        """
        Obtém o histórico de buscas.

        Args:
            limite: Número máximo de entradas

        Returns:
            Lista de buscas do histórico
        """
        return self._history_manager.listar(limite)

    def limpar_historico(self) -> Tuple[bool, str]:
        """
        Limpa o histórico de buscas.

        Returns:
            Tupla (sucesso, mensagem)
        """
        return self._history_manager.limpar()

    def restaurar_busca_historico(self, indice: int) -> Tuple[bool, str, List[SuporteData]]:
        """
        Restaura uma busca do histórico.

        Args:
            indice: Índice da busca no histórico

        Returns:
            Tupla (sucesso, mensagem, resultados)
        """
        historico = self._history_manager.listar()

        if indice < 0 or indice >= len(historico):
            return False, "Índice inválido", []

        entrada = historico[indice]

        # Restaura filtros
        filtros = [self._dict_to_filtro(f) for f in entrada.get('filtros', [])]
        self._filtros_ativos = filtros

        # Executa busca novamente
        resultado = self.buscar(
            texto_geral=entrada.get('texto_geral', ''),
            filtros=filtros,
            salvar_historico=False  # Não salva novamente
        )

        return True, f"Busca de {entrada.get('data_hora', '')} restaurada", resultado

    # === Métodos de Sugestão ===

    def obter_sugestoes_campo(self, campo: str) -> List[Any]:
        """
        Obtém valores únicos para um campo (para autocomplete).

        Args:
            campo: Nome do campo

        Returns:
            Lista de valores únicos ordenada
        """
        suportes = self._repository.listar_todos()

        if campo == 'tag':
            return sorted(set(s.tag for s in suportes))
        elif campo == 'tipo':
            return sorted(set(s.tipo for s in suportes))
        elif campo == 'layer':
            return sorted(set(s.layer for s in suportes))
        else:
            # Propriedade dinâmica
            valores = set()
            for s in suportes:
                if campo in s.propriedades:
                    valores.add(s.propriedades[campo])
            return sorted(valores)

    def obter_campos_disponiveis(self) -> List[Dict[str, str]]:
        """
        Obtém lista de campos disponíveis para filtro.

        Returns:
            Lista de dicts com 'nome' e 'tipo'
        """
        campos = [
            {'nome': 'tag', 'tipo': 'texto', 'label': 'TAG'},
            {'nome': 'tipo', 'tipo': 'texto', 'label': 'Tipo'},
            {'nome': 'layer', 'tipo': 'texto', 'label': 'Camada'}
        ]

        # Adiciona propriedades dinâmicas
        propriedades = self._repository.listar_propriedades_disponiveis()
        for prop in propriedades:
            # Determina tipo baseado em valores
            valores_amostra = []
            for s in self._repository.listar_todos()[:100]:
                if prop in s.propriedades:
                    valores_amostra.append(s.propriedades[prop])

            tipo = 'texto'
            if valores_amostra:
                if all(isinstance(v, (int, float)) for v in valores_amostra if v is not None):
                    tipo = 'numero'

            campos.append({
                'nome': prop,
                'tipo': tipo,
                'label': prop
            })

        return campos
