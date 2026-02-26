"""Conector COM para AutoCAD."""

import time
import pythoncom
import win32com.client
from typing import Any, Dict, List, Optional, Tuple
from dataclasses import dataclass

from .com_error_handler import COMErrorHandler, execute_with_retry


@dataclass
class ConnectionInfo:
    """Informações sobre a conexão com AutoCAD."""
    connected: bool
    application_version: str = ""
    document_name: str = ""
    document_count: int = 0
    model_space_count: int = 0


class AutocadCOMConnector:
    """
    Gerencia a conexão COM com o AutoCAD.

    Fornece métodos para conectar, desconectar e executar operações
    no AutoCAD de forma thread-safe.
    """

    def __init__(self):
        """Inicializa o conector."""
        self._acad: Optional[Any] = None
        self._acad_doc: Optional[Any] = None
        self._acad_model: Optional[Any] = None
        self._is_initialized: bool = False
        self._error_handler = COMErrorHandler()
        # Garante que COM está inicializado na thread atual
        try:
            pythoncom.CoInitialize()
        except Exception:
            # J pode estar inicializado
            pass

    @property
    def is_connected(self) -> bool:
        """Verifica se está conectado ao AutoCAD."""
        return self._acad is not None and self._acad_doc is not None

    @property
    def application(self) -> Optional[Any]:
        """Retorna a aplicação AutoCAD."""
        return self._acad

    @property
    def document(self) -> Optional[Any]:
        """Retorna o documento ativo."""
        return self._acad_doc

    @property
    def model_space(self) -> Optional[Any]:
        """Retorna o ModelSpace do documento ativo."""
        return self._acad_model

    def conectar(self, esperar_documento: bool = True, timeout_seg: int = 30) -> ConnectionInfo:
        """
        Conecta ao AutoCAD.

        Args:
            esperar_documento: Se True, aguarda até haver um documento aberto
            timeout_seg: Tempo máximo de espera por documento

        Returns:
            ConnectionInfo com status da conexão
        """
        try:
            # Tenta conectar ao AutoCAD já aberto
            self._acad = self._try_get_active_object()

            if self._acad is None:
                # Tenta criar nova instância
                self._acad = self._try_create_instance()

            if self._acad is None:
                return ConnectionInfo(connected=False)

            # Aguarda documento ativo
            if esperar_documento:
                if not self._wait_for_document(timeout_seg):
                    return ConnectionInfo(connected=False)

            # Obtém referências ao documento e model space
            self._acad_doc = self._acad.ActiveDocument
            self._acad_model = self._acad_doc.ModelSpace
            self._is_initialized = True

            return ConnectionInfo(
                connected=True,
                application_version=self._acad.Version,
                document_name=self._acad_doc.Name,
                document_count=self._acad.Documents.Count,
                model_space_count=self._acad_model.Count
            )

        except Exception as e:
            self._cleanup()
            return ConnectionInfo(connected=False)

    def _try_get_active_object(self) -> Optional[Any]:
        """Tenta obter instância ativa do AutoCAD."""
        try:
            return win32com.client.GetActiveObject("AutoCAD.Application")
        except Exception:
            return None

    def _try_create_instance(self) -> Optional[Any]:
        """Tenta criar nova instância do AutoCAD."""
        try:
            pythoncom.CoInitialize()
            return win32com.client.Dispatch("AutoCAD.Application")
        except Exception:
            return None

    def _wait_for_document(self, timeout_seg: int) -> bool:
        """
        Aguarda até haver um documento aberto.

        Args:
            timeout_seg: Tempo máximo de espera

        Returns:
            True se obteve documento, False se timeout
        """
        inicio = time.time()

        while self._acad.Documents.Count == 0:
            if time.time() - inicio > timeout_seg:
                return False
            time.sleep(0.5)

        return True

    def _cleanup(self) -> None:
        """Limpa referências COM."""
        self._acad_model = None
        self._acad_doc = None
        self._acad = None
        self._is_initialized = False

    def _ensure_valid_connection(self) -> bool:
        """
        Garante que a conexão COM ainda é válida.
        Tenta reconectar se necessário.

        Returns:
            True se a conexão está válida
        """
        if not self._acad:
            return False

        try:
            # Tenta acessar uma propriedade simples para verificar se ainda está conectado
            _ = self._acad.Version
            return True
        except (pythoncom.com_error, Exception):
            # Objeto COM inválido, tenta reconectar
            self._cleanup()

            # Tenta reconectar sem esperar documento (já deve estar aberto)
            try:
                pythoncom.CoInitialize()
                self._acad = self._try_get_active_object()

                if self._acad and self._acad.Documents.Count > 0:
                    self._acad_doc = self._acad.ActiveDocument
                    self._acad_model = self._acad_doc.ModelSpace
                    self._is_initialized = True
                    return True
            except Exception:
                pass

            return False

    def desconectar(self) -> None:
        """Limpa as referências do AutoCAD (não fecha o aplicativo)."""
        self._cleanup()

    def listar_blocos_suporte(self, tag_atributo: str = "POSICAO", prefixo_bloco: str = "SP_") -> List[Dict[str, Any]]:
        """
        Lista todos os blocos de suporte no ModelSpace.

        Args:
            tag_atributo: Nome do atributo que identifica a posição
            prefixo_bloco: Prefixo do nome do bloco para filtrar (default: "SP_")

        Returns:
            Lista de dicionários com dados dos blocos
        """
        # Garante conexão válida (tentando reconectar se necessário)
        if not self._ensure_valid_connection():
            print("[DEBUG] listar_blocos_suporte: Conexão inválida")
            return []

        blocos = []

        try:
            def collect_blocks():
                result = []
                # No AutoCAD COM, precisamos usar o Count e Item para iterar
                count = self._acad_model.Count
                print(f"[DEBUG] Iterando {count} entidades no ModelSpace (filtro: {prefixo_bloco}*)...")

                # Contadores para estatísticas
                skips = {
                    'not_block_ref': 0,
                    'wrong_prefix': 0,
                    'no_attributes': 0,
                    'no_position_attr': 0,
                    'success': 0,
                    'exceptions': 0
                }

                for i in range(count):
                    # Mostra progresso mais frequente para conjuntos pequenos
                    if i % 10 == 0 or i == count - 1:
                        print(f"[DEBUG] Processando entidade {i}/{count} ({len(result)} suportes encontrados)")

                    try:
                        entity = self._acad_model.Item(i)

                        # FILTRO 1: Primeiro verifica se é BlockReference (skip rápido)
                        if entity.EntityName != 'AcDbBlockReference':
                            skips['not_block_ref'] += 1
                            continue

                        # Debug: mostra blocos encontrados com nome
                        entity_name = entity.Name
                        entity_handle = entity.Handle

                        # FILTRO 2: Verifica prefixo do nome ANTES de verificar atributos
                        if not entity_name.startswith(prefixo_bloco):
                            skips['wrong_prefix'] += 1
                            print(f"[DEBUG] SKIP (prefixo): Handle={entity_handle}, Nome='{entity_name}' não começa com '{prefixo_bloco}'")
                            continue

                        # FILTRO 3: Só então verifica atributos
                        if not entity.HasAttributes:
                            skips['no_attributes'] += 1
                            print(f"[DEBUG] SKIP (sem atributos): Handle={entity_handle}, Nome='{entity_name}' não tem atributos")
                            continue

                        # Busca atributo POSICAO
                        tag_suporte = ""
                        attribs = entity.GetAttributes()
                        for attrib in attribs:
                            if attrib.TagString.upper() == tag_atributo:
                                tag_suporte = attrib.TextString
                                break

                        if not tag_suporte:
                            skips['no_position_attr'] += 1
                            # Lista todos os atributos disponíveis para debug
                            attr_tags = [a.TagString for a in attribs]
                            print(f"[DEBUG] SKIP (sem POSICAO): Handle={entity_handle}, Nome='{entity_name}', Atributos=[{', '.join(attr_tags)}]")
                            continue

                        skips['success'] += 1
                        print(f"[DEBUG] OK: Handle={entity_handle}, Nome='{entity_name}', POSICAO='{tag_suporte}'")

                        insertion_point = entity.InsertionPoint
                        result.append({
                            'tag': tag_suporte,
                            'tipo': entity_name,
                            'handle': entity_handle,
                            'layer': entity.Layer,
                            'posicao_x': float(insertion_point[0]),
                            'posicao_y': float(insertion_point[1]),
                            'posicao_z': float(insertion_point[2]),
                            'is_dynamic': entity.IsDynamicBlock
                        })
                    except Exception as e:
                        skips['exceptions'] += 1
                        print(f"[DEBUG] EXCEÇÃO na entidade {i}: {e}")
                        continue

                # Resumo estatístico
                print(f"[DEBUG] === RESUMO ===")
                print(f"[DEBUG] Total entidades: {count}")
                print(f"[DEBUG] Suportes encontrados: {skips['success']}")
                print(f"[DEBUG] Skips (não é bloco): {skips['not_block_ref']}")
                print(f"[DEBUG] Skips (prefixo errado): {skips['wrong_prefix']}")
                print(f"[DEBUG] Skips (sem atributos): {skips['no_attributes']}")
                print(f"[DEBUG] Skips (sem POSICAO): {skips['no_position_attr']}")
                print(f"[DEBUG] Exceções: {skips['exceptions']}")
                print(f"[DEBUG] Iteração concluída: {len(result)} suportes encontrados")
                return result

            blocos = execute_with_retry(collect_blocks, "Listar blocos de suporte")

        except Exception as e:
            print(f"Erro ao listar blocos: {e}")

        return blocos

    def obter_propriedades_bloco(self, handle: str) -> Dict[str, Any]:
        """
        Obtém propriedades dinâmicas de um bloco pelo handle.

        Args:
            handle: Handle do bloco

        Returns:
            Dicionário de propriedades
        """
        # Garante conexão válida
        if not self._ensure_valid_connection():
            print(f"[DEBUG] obter_propriedades_bloco({handle}): Conexão inválida")
            return {}

        propriedades = {}

        try:
            def get_props():
                print(f"[DEBUG] Buscando propriedades para handle {handle}...")
                count = self._acad_model.Count
                for i in range(count):
                    try:
                        entity = self._acad_model.Item(i)
                        if entity.EntityName == 'AcDbBlockReference' and entity.Handle == handle:
                            if entity.IsDynamicBlock:
                                props = {}
                                dyn_props = entity.GetDynamicBlockProperties()
                                for dyn_prop in dyn_props:
                                    if dyn_prop.PropertyName != "Origin":
                                        valor = dyn_prop.Value
                                        props[dyn_prop.PropertyName] = {
                                            'valor': valor,
                                            'show': dyn_prop.Show,
                                            'readonly': not getattr(dyn_prop, 'ReadOnly', False)
                                        }

                                        # Obtém limites se existirem
                                        if hasattr(dyn_prop, 'ValueMinimum'):
                                            props[dyn_prop.PropertyName]['min'] = dyn_prop.ValueMinimum
                                        if hasattr(dyn_prop, 'ValueMaximum'):
                                            props[dyn_prop.PropertyName]['max'] = dyn_prop.ValueMaximum
                                print(f"[DEBUG] Propriedades encontradas: {len(props)}")
                                return props
                    except Exception:
                        continue
                print(f"[DEBUG] Nenhuma propriedade encontrada para handle {handle}")
                return {}

            propriedades = execute_with_retry(get_props, f"Obter propriedades do bloco {handle}")

        except Exception as e:
            print(f"Erro ao obter propriedades: {e}")

        return propriedades

    def atualizar_propriedade(
        self,
        handle: str,
        nome_propriedade: str,
        novo_valor: Any
    ) -> Tuple[bool, str]:
        """
        Atualiza uma propriedade dinâmica de um bloco.

        Args:
            handle: Handle do bloco
            nome_propriedade: Nome da propriedade
            novo_valor: Novo valor

        Returns:
            Tupla (sucesso, mensagem)
        """
        # Garante conexão válida
        if not self._ensure_valid_connection():
            return False, "Não conectado ao AutoCAD"

        try:
            def update_prop():
                # Itera usando Item method
                count = self._acad_model.Count
                for i in range(count):
                    try:
                        entity = self._acad_model.Item(i)
                        if entity.EntityName == 'AcDbBlockReference' and entity.Handle == handle:
                            dyn_props = entity.GetDynamicBlockProperties()
                            for prop in dyn_props:
                                if prop.PropertyName == nome_propriedade:
                                    # Verifica limites se existirem
                                    if hasattr(prop, 'ValueMinimum') and hasattr(prop, 'ValueMaximum'):
                                        try:
                                            valor_num = float(novo_valor)
                                            if not (prop.ValueMinimum <= valor_num <= prop.ValueMaximum):
                                                return False, (
                                                    f"Valor {novo_valor} fora dos limites "
                                                    f"[{prop.ValueMinimum}, {prop.ValueMaximum}]"
                                                )
                                        except (ValueError, TypeError):
                                            pass

                                    prop.Value = novo_valor
                                    return True, "Propriedade atualizada com sucesso"
                    except Exception:
                        continue
                return False, "Bloco ou propriedade não encontrada"

            return execute_with_retry(update_prop, f"Atualizar propriedade {nome_propriedade}")

        except Exception as e:
            return False, f"Erro ao atualizar: {str(e)}"

    def zoom_para_ponto(self, x: float, y: float, z: float, margem: float = 200) -> bool:
        """
        Faz zoom para um ponto específico.

        Args:
            x: Coordenada X
            y: Coordenada Y
            z: Coordenada Z
            margem: Margem ao redor do ponto

        Returns:
            True se bem-sucedido
        """
        if not self._ensure_valid_connection() or not self._acad:
            return False

        try:
            def do_zoom():
                p1 = self._create_point(x - margem, y + margem, z)
                p2 = self._create_point(x + margem, y - margem, z)
                self._acad.ZoomWindow(p1, p2)
                return True

            return execute_with_retry(do_zoom, "Zoom para ponto")

        except Exception as e:
            print(f"Erro ao fazer zoom: {e}")
            return False

    def _create_point(self, x: float, y: float, z: float) -> Any:
        """Cria um ponto COM do AutoCAD."""
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            (x, y, z)
        )

    def obter_info_documento(self) -> Dict[str, Any]:
        """
        Obtém informações sobre o documento atual.

        Returns:
            Dicionário com informações do documento
        """
        if not self._ensure_valid_connection():
            return {}

        try:
            return {
                'nome': self._acad_doc.Name,
                'path': self._acad_doc.FullName if hasattr(self._acad_doc, 'FullName') else "",
                'bloco_count': self._acad_doc.Blocks.Count if hasattr(self._acad_doc, 'Blocks') else 0,
                'model_space_count': self._acad_model.Count if self._acad_model else 0,
                'saved': self._acad_doc.Saved if hasattr(self._acad_doc, 'Saved') else True
            }
        except Exception as e:
            return {'erro': str(e)}

    def salvar_documento(self) -> Tuple[bool, str]:
        """
        Salva o documento atual.

        Returns:
            Tupla (sucesso, mensagem)
        """
        if not self._ensure_valid_connection() or not self._acad_doc:
            return False, "Não há documento ativo"

        try:
            def do_save():
                self._acad_doc.Save()
                return True, "Documento salvo"

            return execute_with_retry(do_save, "Salvar documento")

        except Exception as e:
            return False, f"Erro ao salvar: {str(e)}"
