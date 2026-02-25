"""Worker thread para carregar suportes do AutoCAD."""

import pythoncom
from typing import List, Optional

from PySide6.QtCore import QThread, Signal

from core.models import SuporteData
from core.repository import SuporteRepository


class LoadSuportesWorker(QThread):
    """
    Worker thread para carregar suportes do AutoCAD.

    Sinais:
        progress: Emitido durante o carregamento (percentual)
        finished: Emitido ao finalizar (lista de suportes)
        error: Emitido em caso de erro (mensagem)
        status: Emitido com status textual
    """

    progress = Signal(int)
    finished = Signal(list)
    error = Signal(str)
    status = Signal(str)

    def __init__(self, parent=None):
        """Inicializa o worker."""
        super().__init__(parent)
        self._repository: Optional[SuporteRepository] = None
        self._carregar_propriedades: bool = True
        self._forcar_recarga: bool = False
        self._cancelado: bool = False

    def configurar(
        self,
        repository: SuporteRepository,
        carregar_propriedades: bool = True,
        forcar_recarga: bool = False
    ) -> None:
        """
        Configura o worker.

        Args:
            repository: Repositório de suportes
            carregar_propriedades: Se True, carrega propriedades dinâmicas
            forcar_recarga: Se True, ignora cache e recarrega
        """
        self._repository = repository
        self._carregar_propriedades = carregar_propriedades
        self._forcar_recarga = forcar_recarga
        self._cancelado = False

    def cancelar(self) -> None:
        """Solicita cancelamento."""
        self._cancelado = True

    def run(self) -> None:
        """Executa o carregamento."""
        try:
            # Inicializa COM nesta thread
            pythoncom.CoInitialize()

            self.status.emit("Conectando ao AutoCAD...")

            if not self._repository.is_connected:
                sucesso, msg = self._repository.conectar()
                if not sucesso:
                    self.error.emit(msg)
                    return

            self.progress.emit(10)
            self.status.emit("Carregando suportes...")

            if self._cancelado:
                self.status.emit("Carregamento cancelado")
                return

            # Carrega suportes básicos
            suportes = self._repository.listar_todos(forcar_recarga=self._forcar_recarga)

            self.progress.emit(50)

            if self._cancelado:
                self.status.emit("Carregamento cancelado")
                return

            # Carrega propriedades dinâmicas se solicitado
            if self._carregar_propriedades:
                total = len(suportes)
                for i, suporte in enumerate(suportes):
                    if self._cancelado:
                        break

                    # Carrega propriedades do bloco
                    props = self._repository.obter_propriedades(suporte.handle)
                    if props:
                        for nome, dados in props.items():
                            if isinstance(dados, dict):
                                suporte.definir_propriedade(nome, dados.get('valor'))
                            else:
                                suporte.definir_propriedade(nome, dados)

                    percentual = 50 + int((i + 1) / total * 50)
                    self.progress.emit(percentual)
                    self.status.emit(f"Carregando propriedades... {i + 1}/{total}")

            if self._cancelado:
                self.status.emit("Carregamento cancelado")
                self.finished.emit([])
                return

            self.progress.emit(100)
            self.status.emit(f"Concluído: {len(suportes)} suportes carregados")
            self.finished.emit(suportes)

        except Exception as e:
            self.error.emit(f"Erro ao carregar suportes: {str(e)}")

        finally:
            # NOTA: Não chamamos CoUninitialize() aqui porque os objetos COM
            # criados (AutoCAD Application) serão usados por outras threads.
            # O COM será limpo automaticamente quando a aplicação terminar.
            pass


class AutoConnectWorker(QThread):
    """
    Worker thread para conexão automática com AutoCAD.

    Sinais:
        connected: Emitido quando conectado (info_dict)
        failed: Emitido quando falha (mensagem)
        status: Emitido com status textual
    """

    connected = Signal(dict)
    failed = Signal(str)
    status = Signal(str)

    def __init__(self, parent=None):
        """Inicializa o worker."""
        super().__init__(parent)
        self._repository: Optional[SuporteRepository] = None
        self._timeout_seg: int = 30

    def configurar(self, repository: SuporteRepository, timeout_seg: int = 30) -> None:
        """
        Configura o worker.

        Args:
            repository: Repositório de suportes
            timeout_seg: Tempo máximo de espera por documento
        """
        self._repository = repository
        self._timeout_seg = timeout_seg

    def run(self) -> None:
        """Executa a conexão."""
        try:
            # Inicializa COM nesta thread
            pythoncom.CoInitialize()

            self.status.emit("Tentando conectar ao AutoCAD...")

            sucesso, msg = self._repository.conectar(
                esperar_documento=True,
                timeout_seg=self._timeout_seg
            )

            if sucesso:
                self.status.emit("Conectado ao AutoCAD")
                info = self._repository.obter_info_documento() if hasattr(self._repository, 'obter_info_documento') else {}
                self.connected.emit(info)
            else:
                self.status.emit("Falha na conexão")
                self.failed.emit(msg)

        except Exception as e:
            self.failed.emit(f"Erro na conexão: {str(e)}")

        finally:
            # NOTA: Não chamamos CoUninitialize() aqui porque os objetos COM
            # criados (AutoCAD Application) serão usados por outras threads.
            # O COM será limpo automaticamente quando a aplicação terminar.
            pass
