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
            print("[DEBUG] LoadSuportesWorker: Iniciando carregamento")

            self.status.emit("Conectando ao AutoCAD...")
            print("[DEBUG] Verificando conexão...")

            if not self._repository.is_connected:
                print("[DEBUG] Não conectado, tentando conectar...")
                sucesso, msg = self._repository.conectar()
                print(f"[DEBUG] Resultado conexão: {sucesso} - {msg}")
                if not sucesso:
                    self.error.emit(msg)
                    return
            else:
                print("[DEBUG] Já conectado ao AutoCAD")

            self.progress.emit(10)
            self.status.emit("Carregando suportes...")

            if self._cancelado:
                self.status.emit("Carregamento cancelado")
                return

            # Carrega suportes básicos
            print("[DEBUG] Chamando listar_todos()...")
            suportes = self._repository.listar_todos(forcar_recarga=self._forcar_recarga)
            print(f"[DEBUG] listar_todos() retornou {len(suportes)} suportes")

            self.progress.emit(50)

            if self._cancelado:
                self.status.emit("Carregamento cancelado")
                return

            # Carrega propriedades dinâmicas se solicitado
            if self._carregar_propriedades:
                total = len(suportes)
                print(f"[DEBUG] Carregando propriedades de {total} suportes...")
                for i, suporte in enumerate(suportes):
                    if self._cancelado:
                        break

                    # Progresso a cada 10 suportes
                    if i % 10 == 0:
                        print(f"[DEBUG] Progresso propriedades: {i+1}/{total}")

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
            print(f"[DEBUG] Carregamento concluído: {len(suportes)} suportes")
            self.status.emit(f"Concluído: {len(suportes)} suportes carregados")
            self.finished.emit(suportes)

        except Exception as e:
            print(f"[DEBUG] Erro no LoadSuportesWorker: {e}")
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
            print("[DEBUG] AutoConnectWorker: Tentando conectar...")

            self.status.emit("Tentando conectar ao AutoCAD...")

            sucesso, msg = self._repository.conectar(
                esperar_documento=True,
                timeout_seg=self._timeout_seg
            )

            print(f"[DEBUG] AutoConnectWorker: Resultado={sucesso}, Msg={msg}")

            if sucesso:
                self.status.emit("Conectado ao AutoCAD")
                info = self._repository.obter_info_documento() if hasattr(self._repository, 'obter_info_documento') else {}
                print(f"[DEBUG] AutoConnectWorker: Info documento={info}")
                self.connected.emit(info)
            else:
                self.status.emit("Falha na conexão")
                self.failed.emit(msg)

        except Exception as e:
            print(f"[DEBUG] AutoConnectWorker: Erro na conexão: {e}")
            self.failed.emit(f"Erro na conexão: {str(e)}")

        finally:
            # NOTA: Não chamamos CoUninitialize() aqui porque os objetos COM
            # criados (AutoCAD Application) serão usados por outras threads.
            # O COM será limpo automaticamente quando a aplicação terminar.
            pass
