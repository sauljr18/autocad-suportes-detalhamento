"""Repositório de suportes AutoCAD."""

from typing import Any, Dict, List, Optional, Tuple

from core.models import SuporteData, FiltroBusca
from utils.autocad_connector import AutocadCOMConnector


class SuporteRepository:
    """
    Abstrai o acesso aos dados de suportes no AutoCAD.

    Fornece uma interface de alto nível para operações de CRUD
    sobre suportes, usando o AutocadCOMConnector internamente.
    """

    def __init__(self):
        """Inicializa o repositório."""
        self._connector = AutocadCOMConnector()
        self._cache: List[SuporteData] = []
        self._cache_dirty = True

    @property
    def is_connected(self) -> bool:
        """Verifica se está conectado ao AutoCAD."""
        return self._connector.is_connected

    def conectar(self, esperar_documento: bool = True, timeout_seg: int = 30) -> Tuple[bool, str]:
        """
        Conecta ao AutoCAD.

        Args:
            esperar_documento: Se True, aguarda documento aberto
            timeout_seg: Tempo máximo de espera

        Returns:
            Tupla (sucesso, mensagem)
        """
        info = self._connector.conectar(esperar_documento, timeout_seg)

        if info.connected:
            return True, f"Conectado ao AutoCAD {info.application_version}"
        else:
            return False, "Não foi possível conectar ao AutoCAD"

    def desconectar(self) -> None:
        """Desconecta do AutoCAD."""
        self._connector.desconectar()
        self._cache.clear()
        self._cache_dirty = True

    def obter_info_documento(self) -> Dict[str, Any]:
        """
        Obtém informações sobre o documento AutoCAD atual.

        Returns:
            Dicionário com informações do documento
        """
        print("[DEBUG] obter_info_documento: Obtendo informações do documento")
        info = self._connector.obter_info_documento()
        print(f"[DEBUG] obter_info_documento: {info}")
        return info

    def listar_todos(self, forcar_recarga: bool = False) -> List[SuporteData]:
        """
        Lista todos os suportes do AutoCAD.

        Args:
            forcar_recarga: Se True, ignora o cache e recarrega

        Returns:
            Lista de SuporteData
        """
        print(f"[DEBUG] listar_todos: forcar_recarga={forcar_recarga}, cache_dirty={self._cache_dirty}")

        if not self.is_connected:
            print("[DEBUG] listar_todos: Não conectado")
            return []

        if not self._cache_dirty and not forcar_recarga:
            print(f"[DEBUG] listar_todos: Retornando cache ({len(self._cache)} suportes)")
            return self._cache.copy()

        print("[DEBUG] listar_todos: Cache inválido, recarregando do AutoCAD...")
        blocos = self._connector.listar_blocos_suporte()
        print(f"[DEBUG] listar_todos: {len(blocos)} blocos retornados pelo conector")

        self._cache.clear()
        for bloco in blocos:
            # Carrega propriedades dinâmicas se for bloco dinâmico
            propriedades = {}
            if bloco.get('is_dynamic'):
                propriedades_raw = self._connector.obter_propriedades_bloco(bloco['handle'])
                # Extrai apenas os valores
                propriedades = {
                    k: v.get('valor', v) if isinstance(v, dict) else v
                    for k, v in propriedades_raw.items()
                }

            suporte = SuporteData(
                tag=bloco['tag'],
                tipo=bloco['tipo'],
                posicao_x=bloco['posicao_x'],
                posicao_y=bloco['posicao_y'],
                posicao_z=bloco['posicao_z'],
                handle=bloco['handle'],
                propriedades=propriedades,
                layer=bloco.get('layer', '')
            )
            self._cache.append(suporte)

        self._cache_dirty = False
        print(f"[DEBUG] listar_todos: {len(self._cache)} suportes no cache")
        return self._cache.copy()

    def buscar_por_filtro(self, filtros: List[FiltroBusca]) -> List[SuporteData]:
        """
        Busca suportes que atendem aos filtros.

        Args:
            filtros: Lista de filtros a aplicar

        Returns:
            Lista de SuporteData filtrada
        """
        suportes = self.listar_todos()

        if not filtros:
            return suportes

        resultado = suportes.copy()

        for filtro in filtros:
            resultado = [s for s in resultado if filtro.verificar(s)]

        return resultado

    def buscar_por_tag(self, tag: str) -> Optional[SuporteData]:
        """
        Busca um suporte pela tag.

        Args:
            tag: Tag do suporte

        Returns:
            SuporteData ou None se não encontrado
        """
        suportes = self.listar_todos()

        for suporte in suportes:
            if suporte.tag.upper() == tag.upper():
                return suporte

        return None

    def buscar_por_handle(self, handle: str) -> Optional[SuporteData]:
        """
        Busca um suporte pelo handle.

        Args:
            handle: Handle do suporte

        Returns:
            SuporteData ou None se não encontrado
        """
        suportes = self.listar_todos()

        for suporte in suportes:
            if suporte.handle == handle:
                return suporte

        return None

    def obter_propriedades(self, handle: str) -> Dict[str, Any]:
        """
        Obtém propriedades dinâmicas de um suporte.

        Args:
            handle: Handle do suporte

        Returns:
            Dicionário de propriedades
        """
        if not self.is_connected:
            return {}

        props_raw = self._connector.obter_propriedades_bloco(handle)

        # Converte para formato simplificado
        propriedades = {}
        for nome, dados in props_raw.items():
            if isinstance(dados, dict):
                propriedades[nome] = {
                    'valor': dados.get('valor'),
                    'min': dados.get('min'),
                    'max': dados.get('max'),
                    'readonly': dados.get('readonly', False)
                }
            else:
                propriedades[nome] = {'valor': dados}

        return propriedades

    def atualizar_propriedade(
        self,
        handle: str,
        propriedade: str,
        valor: Any
    ) -> Tuple[bool, str]:
        """
        Atualiza uma propriedade de um suporte.

        Args:
            handle: Handle do suporte
            propriedade: Nome da propriedade
            valor: Novo valor

        Returns:
            Tupla (sucesso, mensagem)
        """
        if not self.is_connected:
            return False, "Não conectado ao AutoCAD"

        sucesso, mensagem = self._connector.atualizar_propriedade(handle, propriedade, valor)

        if sucesso:
            # Atualiza o cache se existir
            for suporte in self._cache:
                if suporte.handle == handle:
                    suporte.definir_propriedade(propriedade, valor)
                    break

        return sucesso, mensagem

    def atualizar_lote(
        self,
        handles: List[str],
        propriedade: str,
        valor: Any
    ) -> Dict[str, Any]:
        """
        Atualiza uma propriedade em múltiplos suportes.

        Args:
            handles: Lista de handles dos suportes
            propriedade: Nome da propriedade
            valor: Novo valor

        Returns:
            Dicionário com estatísticas:
            {
                'total': int,
                'sucesso': int,
                'falhas': int,
                'detalhes': List[Dict]
            }
        """
        stats = {
            'total': len(handles),
            'sucesso': 0,
            'falhas': 0,
            'detalhes': []
        }

        if not self.is_connected:
            stats['falhas'] = len(handles)
            stats['detalhes'] = [
                {'handle': h, 'erro': 'Não conectado ao AutoCAD'}
                for h in handles
            ]
            return stats

        for handle in handles:
            sucesso, mensagem = self.atualizar_propriedade(handle, propriedade, valor)

            if sucesso:
                stats['sucesso'] += 1
                stats['detalhes'].append({
                    'handle': handle,
                    'sucesso': True,
                    'mensagem': mensagem
                })
            else:
                stats['falhas'] += 1
                stats['detalhes'].append({
                    'handle': handle,
                    'sucesso': False,
                    'erro': mensagem
                })

        return stats

    def zoom_para_suporte(self, handle: str, margem: float = 200) -> Tuple[bool, str]:
        """
        Faz zoom para um suporte.

        Args:
            handle: Handle do suporte
            margem: Margem ao redor do ponto

        Returns:
            Tupla (sucesso, mensagem)
        """
        if not self.is_connected:
            return False, "Não conectado ao AutoCAD"

        # Busca o suporte para obter coordenadas
        suporte = self.buscar_por_handle(handle)

        if not suporte:
            return False, f"Suporte com handle {handle} não encontrado"

        sucesso = self._connector.zoom_para_ponto(
            suporte.posicao_x,
            suporte.posicao_y,
            suporte.posicao_z,
            margem
        )

        if sucesso:
            return True, f"Zoom para {suporte.tag}"
        else:
            return False, "Falha ao executar zoom"

    def listar_tipos_suporte(self) -> List[str]:
        """
        Lista todos os tipos de suporte únicos.

        Returns:
            Lista de tipos
        """
        suportes = self.listar_todos()
        tipos = sorted(set(s.tipo for s in suportes))
        return tipos

    def listar_camadas(self) -> List[str]:
        """
        Lista todas as camadas únicas.

        Returns:
            Lista de camadas
        """
        suportes = self.listar_todos()
        camadas = sorted(set(s.layer for s in suportes if s.layer))
        return camadas

    def listar_propriedades_disponiveis(self) -> List[str]:
        """
        Lista todas as propriedades dinâmicas disponíveis.

        Returns:
            Lista de nomes de propriedades
        """
        suportes = self.listar_todos()
        propriedades = set()

        for suporte in suportes:
            propriedades.update(suporte.listar_nomes_propriedades())

        return sorted(propriedades)

    def obter_estatisticas(self) -> Dict[str, Any]:
        """
        Obtém estatísticas sobre os suportes carregados.

        Returns:
            Dicionário com estatísticas
        """
        suportes = self.listar_todos()

        if not suportes:
            return {
                'total': 0,
                'tipos': {},
                'camadas': {},
                'media_posicao': {'x': 0, 'y': 0, 'z': 0}
            }

        tipos = {}
        camadas = {}

        for s in suportes:
            tipos[s.tipo] = tipos.get(s.tipo, 0) + 1
            camadas[s.layer] = camadas.get(s.layer, 0) + 1

        return {
            'total': len(suportes),
            'tipos': tipos,
            'camadas': camadas,
            'media_posicao': {
                'x': sum(s.posicao_x for s in suportes) / len(suportes),
                'y': sum(s.posicao_y for s in suportes) / len(suportes),
                'z': sum(s.posicao_z for s in suportes) / len(suportes)
            }
        }
