"""Worker thread para edição em lote de suportes."""

import pythoncom
from typing import Any, Dict, List

from PySide6.QtCore import QThread, Signal

from core.models import SuporteData
from core.repository import SuporteRepository


class BatchEditWorker(QThread):
    """
    Worker thread para edição em lote de propriedades.

    Sinais:
        progress: Emitido durante o processamento (atual, total)
        finished: Emitido ao finalizar (estatísticas)
        error: Emitido em caso de erro (mensagem)
        status: Emitido com status textual
    """

    progress = Signal(int, int)
    finished = Signal(dict)
    error = Signal(str)
    status = Signal(str)

    def __init__(self, parent=None):
        """Inicializa o worker."""
        super().__init__(parent)
        self._repository: SuporteRepository = None
        self._suportes: List[SuporteData] = []
        self._propriedade: str = ""
        self._valor: Any = None
        self._cancelado: bool = False

    def configurar(
        self,
        repository: SuporteRepository,
        suportes: List[SuporteData],
        propriedade: str,
        valor: Any
    ) -> None:
        """
        Configura o worker.

        Args:
            repository: Repositório de suportes
            suportes: Lista de suportes a editar
            propriedade: Nome da propriedade a editar
            valor: Novo valor
        """
        self._repository = repository
        self._suportes = suportes
        self._propriedade = propriedade
        self._valor = valor
        self._cancelado = False

    def cancelar(self) -> None:
        """Solicita cancelamento."""
        self._cancelado = True

    def run(self) -> None:
        """Executa a edição em lote."""
        try:
            # Inicializa COM nesta thread
            pythoncom.CoInitialize()

            total = len(self._suportes)
            if total == 0:
                self.error.emit("Nenhum suporte selecionado")
                return

            self.status.emit(f"Iniciando edição em lote: {total} suportes")

            estatisticas = {
                'total': total,
                'sucesso': 0,
                'falhas': 0,
                'detalhes': []
            }

            for i, suporte in enumerate(self._suportes):
                if self._cancelado:
                    self.status.emit("Edição cancelada")
                    break

                self.status.emit(
                    f"Editando {suporte.tag} ({i + 1}/{total})..."
                )
                self.progress.emit(i + 1, total)

                sucesso, mensagem = self._repository.atualizar_propriedade(
                    suporte.handle,
                    self._propriedade,
                    self._valor
                )

                if sucesso:
                    estatisticas['sucesso'] += 1
                    estatisticas['detalhes'].append({
                        'tag': suporte.tag,
                        'handle': suporte.handle,
                        'sucesso': True,
                        'mensagem': mensagem
                    })
                else:
                    estatisticas['falhas'] += 1
                    estatisticas['detalhes'].append({
                        'tag': suporte.tag,
                        'handle': suporte.handle,
                        'sucesso': False,
                        'erro': mensagem
                    })

            self.progress.emit(total, total)
            self.finished.emit(estatisticas)

        except Exception as e:
            self.error.emit(f"Erro na edição em lote: {str(e)}")

        finally:
            # Limpa COM
            pythoncom.CoUninitialize()


class MultiPropertyEditWorker(QThread):
    """
    Worker thread para edição de múltiplas propriedades em múltiplos suportes.

    Sinais:
        progress: Emitido durante o processamento (atual, total)
        finished: Emitido ao finalizar (estatísticas)
        error: Emitido em caso de erro (mensagem)
        status: Emitido com status textual
    """

    progress = Signal(int, int)
    finished = Signal(dict)
    error = Signal(str)
    status = Signal(str)

    def __init__(self, parent=None):
        """Inicializa o worker."""
        super().__init__(parent)
        self._repository: SuporteRepository = None
        self._suportes: List[SuporteData] = []
        self._propriedades: Dict[str, Any] = {}
        self._cancelado: bool = False

    def configurar(
        self,
        repository: SuporteRepository,
        suportes: List[SuporteData],
        propriedades: Dict[str, Any]
    ) -> None:
        """
        Configura o worker.

        Args:
            repository: Repositório de suportes
            suportes: Lista de suportes a editar
            propriedades: Dicionário de propriedades e valores
        """
        self._repository = repository
        self._suportes = suportes
        self._propriedades = propriedades
        self._cancelado = False

    def cancelar(self) -> None:
        """Solicita cancelamento."""
        self._cancelado = True

    def run(self) -> None:
        """Executa a edição multi-propriedade."""
        try:
            # Inicializa COM nesta thread
            pythoncom.CoInitialize()

            total_ops = len(self._suportes) * len(self._propriedades)
            if total_ops == 0:
                self.error.emit("Nenhuma operação a executar")
                return

            self.status.emit(
                f"Iniciando edição: {len(self._suportes)} suportes, "
                f"{len(self._propriedades)} propriedades"
            )

            estatisticas = {
                'total_suportes': len(self._suportes),
                'total_propriedades': len(self._propriedades),
                'total_operacoes': total_ops,
                'sucesso': 0,
                'falhas': 0,
                'por_suporte': []
            }

            operacao = 0

            for suporte in self._suportes:
                if self._cancelado:
                    self.status.emit("Edição cancelada")
                    break

                stats_suporte = {
                    'tag': suporte.tag,
                    'handle': suporte.handle,
                    'propriedades': {}
                }

                for prop_nome, prop_valor in self._propriedades.items():
                    if self._cancelado:
                        break

                    operacao += 1
                    self.status.emit(
                        f"Editando {suporte.tag}: {prop_nome} ({operacao}/{total_ops})"
                    )
                    self.progress.emit(operacao, total_ops)

                    sucesso, mensagem = self._repository.atualizar_propriedade(
                        suporte.handle,
                        prop_nome,
                        prop_valor
                    )

                    stats_suporte['propriedades'][prop_nome] = {
                        'sucesso': sucesso,
                        'mensagem': mensagem
                    }

                    if sucesso:
                        estatisticas['sucesso'] += 1
                    else:
                        estatisticas['falhas'] += 1

                estatisticas['por_suporte'].append(stats_suporte)

            self.progress.emit(total_ops, total_ops)
            self.finished.emit(estatisticas)

        except Exception as e:
            self.error.emit(f"Erro na edição: {str(e)}")

        finally:
            # Limpa COM
            pythoncom.CoUninitialize()
