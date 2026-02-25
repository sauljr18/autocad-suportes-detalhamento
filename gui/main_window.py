"""Janela principal da aplicação de navegação de suportes."""

import os
import sys
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QAction, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QDockWidget, QHBoxLayout, QLabel, QMainWindow,
    QMessageBox, QProgressBar, QSplitter, QStatusBar, QVBoxLayout,
    QWidget, QToolBar
)

from core.models import SuporteData, FiltroBusca
from core.repository import SuporteRepository
from services.search_service import SearchService

from gui.panels.search_panel import SearchPanel
from gui.panels.table_panel import TablePanel
from gui.panels.edit_panel import EditPanel
from gui.models.suporte_table_model import SuporteTableModel

from workers.autocad_worker import LoadSuportesWorker, AutoConnectWorker
from workers.batch_edit_worker import BatchEditWorker


class MainWindow(QMainWindow):
    """Janela principal da aplicação."""

    def __init__(self):
        super().__init__()

        self._repository = SuporteRepository()
        self._search_service: Optional[SearchService] = None
        self._suportes_carregados: List[SuporteData] = []

        self._load_worker: Optional[LoadSuportesWorker] = None
        self._batch_worker: Optional[BatchEditWorker] = None

        self._setup_ui()
        self._criar_menu()
        self._criar_toolbar()
        self._criar_status_bar()

        # Timer para status de conexão
        self._status_timer = QTimer(self)
        self._status_timer.timeout.connect(self._atualizar_status_conexao)
        self._status_timer.start(5000)

        # Tenta conectar automaticamente ao iniciar
        QTimer.singleShot(500, self._conectar_autocad)

    def _setup_ui(self) -> None:
        """Configura a UI principal."""
        self.setWindowTitle("Navegação de Suportes AutoCAD")
        self.setGeometry(100, 100, 1400, 900)

        # Widget central
        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)

        # Splitter principal
        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        # Painel esquerdo (busca e tabela)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Painel de busca
        self._search_panel = SearchPanel()
        self._search_panel.busca_solicitada.connect(self._on_busca)
        self._search_panel.filtro_adicionado.connect(self._on_filtro_adicionado)
        self._search_panel.filtro_removido.connect(self._on_filtro_removido)
        self._search_panel.limpar_solicitado.connect(self._on_limpar)
        self._search_panel.preset_carregado.connect(self._carregar_preset)
        self._search_panel.preset_salvo.connect(self._salvar_preset)
        self._search_panel.preset_gerenciar.connect(self._gerenciar_presets)
        self._search_panel.historico_navegar.connect(self._navegar_historico)

        left_layout.addWidget(self._search_panel)

        # Painel de tabela
        self._table_panel = TablePanel()
        self._table_panel.suporte_selecionado.connect(self._on_suporte_selecionado)
        self._table_panel.zoom_solicitado.connect(self._on_zoom)
        self._table_panel.editar_solicitado.connect(self._on_editar)
        self._table_panel.selecao_mudou.connect(self._on_selecao_mudou)
        self._table_panel.atualizar_solicitado.connect(self._atualizar_dados)

        left_layout.addWidget(self._table_panel)

        splitter.addWidget(left_widget)

        # Painel direito (edição)
        self._edit_panel = EditPanel()
        self._edit_panel.valor_alterado.connect(self._on_valor_alterado)
        self._edit_panel.aplicar_lote.connect(self._on_aplicar_lote)

        # Dock ou widget fixo para edição
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.addWidget(self._edit_panel)

        splitter.addWidget(right_widget)

        # Proporção inicial
        splitter.setSizes([900, 400])

    def _criar_menu(self) -> None:
        """Cria a barra de menu."""
        menubar = self.menuBar()

        # Menu Arquivo
        menu_arquivo = menubar.addMenu("&Arquivo")

        action_conectar = QAction("&Conectar AutoCAD", self)
        action_conectar.setShortcut(QKeySequence("Ctrl+C"))
        action_conectar.triggered.connect(self._conectar_autocad)
        menu_arquivo.addAction(action_conectar)

        action_desconectar = QAction("&Desconectar", self)
        action_desconectar.triggered.connect(self._desconectar_autocad)
        menu_arquivo.addAction(action_desconectar)

        menu_arquivo.addSeparator()

        action_sair = QAction("&Sair", self)
        action_sair.setShortcut(QKeySequence("Ctrl+Q"))
        action_sair.triggered.connect(self.close)
        menu_arquivo.addAction(action_sair)

        # Menu Editar
        menu_editar = menubar.addMenu("&Editar")

        action_atualizar = QAction("&Atualizar Dados", self)
        action_atualizar.setShortcut(QKeySequence("F5"))
        action_atualizar.triggered.connect(self._atualizar_dados)
        menu_editar.addAction(action_atualizar)

        action_selecionar_todos = QAction("&Selecionar Todos", self)
        action_selecionar_todos.setShortcut(QKeySequence("Ctrl+A"))
        action_selecionar_todos.triggered.connect(self._table_panel._selecionar_todos)
        menu_editar.addAction(action_selecionar_todos)

        # Menu Visualizar
        menu_visualizar = menubar.addMenu("&Visualizar")

        action_ordenar_tag = QAction("Ordenar por &TAG", self)
        action_ordenar_tag.triggered.connect(self._ordenar_por_tag)
        menu_visualizar.addAction(action_ordenar_tag)

        action_ordenar_tipo = QAction("Ordenar por &Tipo", self)
        action_ordenar_tipo.triggered.connect(self._ordenar_por_tipo)
        menu_visualizar.addAction(action_ordenar_tipo)

        # Menu Ferramentas
        menu_ferramentas = menubar.addMenu("&Ferramentas")

        action_zoom_selecionado = QAction("Zoom para Selecionado", self)
        action_zoom_selecionado.setShortcut(QKeySequence("Ctrl+Z"))
        action_zoom_selecionado.triggered.connect(self._zoom_para_selecionado)
        menu_ferramentas.addAction(action_zoom_selecionado)

        # Menu Ajuda
        menu_ajuda = menubar.addMenu("&Ajuda")

        action_sobre = QAction("&Sobre", self)
        action_sobre.triggered.connect(self._mostrar_sobre)
        menu_ajuda.addAction(action_sobre)

    def _criar_toolbar(self) -> None:
        """Cria a barra de ferramentas."""
        toolbar = QToolBar("Principal", self)
        self.addToolBar(toolbar)

        action_conectar = QAction("🔌 Conectar", self)
        action_conectar.setToolTip("Conectar ao AutoCAD")
        action_conectar.triggered.connect(self._conectar_autocad)
        toolbar.addAction(action_conectar)

        action_atualizar = QAction("🔄 Atualizar", self)
        action_atualizar.setToolTip("Atualizar dados do AutoCAD")
        action_atualizar.triggered.connect(self._atualizar_dados)
        toolbar.addAction(action_atualizar)

        toolbar.addSeparator()

        action_zoom = QAction("🔍 Zoom", self)
        action_zoom.setToolTip("Zoom para suporte selecionado")
        action_zoom.triggered.connect(self._zoom_para_selecionado)
        toolbar.addAction(action_zoom)

        action_editar = QAction("✏️ Editar", self)
        action_editar.setToolTip("Editar propriedades")
        action_editar.triggered.connect(self._on_editar_selecionado)
        toolbar.addAction(action_editar)

    def _criar_status_bar(self) -> None:
        """Cria a barra de status."""
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)

        # Label de conexão
        self._status_conexao = QLabel("● Desconectado")
        self._status_conexao.setStyleSheet("color: red;")
        self._status_bar.addWidget(self._status_conexao)

        # Separator
        self._status_bar.addPermanentWidget(QLabel("|"))

        # Label de contagem
        self._status_contagem = QLabel("Suportes: 0")
        self._status_bar.addPermanentWidget(self._status_contagem)

        # Separator
        self._status_bar.addPermanentWidget(QLabel("|"))

        # Label de filtrados
        self._status_filtrados = QLabel("Filtrados: 0")
        self._status_bar.addPermanentWidget(self._status_filtrados)

        # Separator
        self._status_bar.addPermanentWidget(QLabel("|"))

        # Label de selecionados
        self._status_selecionados = QLabel("Selecionados: 0")
        self._status_bar.addPermanentWidget(self._status_selecionados)

        # Progress bar
        self._progress_bar = QProgressBar()
        self._progress_bar.setVisible(False)
        self._progress_bar.setMaximumWidth(150)
        self._status_bar.addPermanentWidget(self._progress_bar)

    # === Conexão AutoCAD ===

    def _conectar_autocad(self) -> None:
        """Conecta ao AutoCAD."""
        if self._repository.is_connected:
            QMessageBox.information(self, "Info", "Já conectado ao AutoCAD")
            return

        self._mostrar_status("Conectando ao AutoCAD...")

        worker = AutoConnectWorker(self)
        worker.configurar(self._repository, timeout_seg=30)
        worker.connected.connect(self._on_conectado)
        worker.failed.connect(self._on_conexao_falhou)
        worker.status.connect(self._mostrar_status)
        worker.start()

    def _on_conectado(self, info: Dict[str, Any]) -> None:
        """Trata conexão bem-sucedida."""
        self._status_conexao.setText("● Conectado")
        self._status_conexao.setStyleSheet("color: green;")
        self._mostrar_status(f"Conectado: {info.get('nome', 'AutoCAD')}")

        # Inicializa serviço de busca
        if self._search_service is None:
            self._search_service = SearchService(self._repository)

        # Atualiza campos de filtro
        self._atualizar_campos_filtro()

        # Carrega dados automaticamente
        self._atualizar_dados()

    def _on_conexao_falhou(self, mensagem: str) -> None:
        """Trata falha de conexão."""
        self._status_conexao.setText("● Desconectado")
        self._status_conexao.setStyleSheet("color: red;")
        QMessageBox.warning(self, "Erro de Conexão", mensagem)

    def _desconectar_autocad(self) -> None:
        """Desconecta do AutoCAD."""
        self._repository.desconectar()
        self._status_conexao.setText("● Desconectado")
        self._status_conexao.setStyleSheet("color: red;")
        self._table_panel.limpar()
        self._edit_panel.limpar()

    def _atualizar_status_conexao(self) -> None:
        """Atualiza status de conexão periodicamente."""
        if self._repository.is_connected:
            self._status_conexao.setText("● Conectado")
            self._status_conexao.setStyleSheet("color: green;")
        else:
            self._status_conexao.setText("● Desconectado")
            self._status_conexao.setStyleSheet("color: red;")

    # === Carregamento de Dados ===

    def _atualizar_dados(self) -> None:
        """Atualiza os dados do AutoCAD."""
        if not self._repository.is_connected:
            QMessageBox.warning(self, "Aviso", "Não conectado ao AutoCAD")
            return

        if self._load_worker and self._load_worker.isRunning():
            return

        self._progress_bar.setVisible(True)
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)

        self._load_worker = LoadSuportesWorker(self)
        self._load_worker.configurar(self._repository, carregar_propriedades=True)
        self._load_worker.progress.connect(self._progress_bar.setValue)
        self._load_worker.finished.connect(self._on_dados_carregados)
        self._load_worker.error.connect(self._on_erro_carregamento)
        self._load_worker.status.connect(self._mostrar_status)
        self._load_worker.start()

    def _on_dados_carregados(self, suportes: List[SuporteData]) -> None:
        """Trata dados carregados."""
        self._suportes_carregados = suportes
        self._table_panel.atualizar_dados(suportes)

        self._status_contagem.setText(f"Suportes: {len(suportes)}")
        self._status_filtrados.setText(f"Filtrados: {len(suportes)}")

        self._progress_bar.setVisible(False)
        self._progress_bar.setValue(0)

        self._mostrar_status(f"{len(suportes)} suportes carregados")

    def _on_erro_carregamento(self, mensagem: str) -> None:
        """Trata erro no carregamento."""
        self._progress_bar.setVisible(False)
        QMessageBox.critical(self, "Erro", mensagem)

    # === Busca e Filtros ===

    def _on_busca(self, texto: str, filtros: List[FiltroBusca]) -> None:
        """Trata solicitação de busca."""
        if self._search_service is None:
            return

        resultado = self._search_service.buscar(texto_geral=texto, filtros=filtros)
        self._table_panel.atualizar_dados(resultado)

        self._status_filtrados.setText(f"Filtrados: {len(resultado)}")

    def _on_filtro_adicionado(self, filtro: FiltroBusca) -> None:
        """Trata filtro adicionado."""
        # A busca é tratada pelo sinal busca_solicitada
        pass

    def _on_filtro_removido(self, indice: int) -> None:
        """Trata filtro removido."""
        # Refaz busca sem o filtro removido
        texto = self._search_panel.texto_busca
        filtros = self._search_panel.filtros_ativos
        self._on_busca(texto, filtros)

    def _on_limpar(self) -> None:
        """Trata solicitação de limpar busca."""
        if self._search_service is None:
            return

        self._search_service.limpar_filtros()
        self._table_panel.atualizar_dados(self._suportes_carregados)
        self._status_filtrados.setText(f"Filtrados: {len(self._suportes_carregados)}")

    def _atualizar_campos_filtro(self) -> None:
        """Atualiza campos disponíveis para filtro."""
        if self._search_service is None:
            return

        campos = self._search_service.obter_campos_disponiveis()
        self._search_panel.definir_campos(campos)

    # === Seleção e Edição ===

    def _on_suporte_selecionado(self, suporte: SuporteData) -> None:
        """Trata suporte selecionado."""
        self._edit_panel.definir_suporte(suporte)

    def _on_selecao_mudou(self, qtd: int) -> None:
        """Trata mudança na quantidade de selecionados."""
        self._status_selecionados.setText(f"Selecionados: {qtd}")

    def _on_zoom(self, handle: str) -> None:
        """Trata solicitação de zoom."""
        if not self._repository.is_connected:
            return

        sucesso, mensagem = self._repository.zoom_para_suporte(handle)
        if sucesso:
            self._mostrar_status(mensagem)
        else:
            QMessageBox.warning(self, "Aviso", mensagem)

    def _zoom_para_selecionado(self) -> None:
        """Faz zoom para o suporte selecionado."""
        suporte = self._table_panel.obter_suporte_selecionado()
        if suporte:
            self._on_zoom(suporte.handle)
        else:
            QMessageBox.information(self, "Info", "Selecione um suporte")

    def _on_editar(self, handle: str) -> None:
        """Trata solicitação de edição."""
        suporte = self._repository.buscar_por_handle(handle)
        if suporte:
            self._edit_panel.definir_suporte(suporte)
            self._table_panel.selecionar_por_handle(handle)

    def _on_editar_selecionado(self) -> None:
        """Edita o suporte selecionado."""
        suporte = self._table_panel.obter_suporte_selecionado()
        if suporte:
            self._on_editar(suporte.handle)

    def _on_valor_alterado(self, handle: str, propriedade: str, valor: Any) -> None:
        """Trata valor alterado."""
        if not self._repository.is_connected:
            return

        sucesso, mensagem = self._repository.atualizar_propriedade(handle, propriedade, valor)

        if sucesso:
            self._mostrar_status(f"Propriedade '{propriedade}' atualizada")
            # Recarrega o suporte
            self._on_editar(handle)
        else:
            QMessageBox.warning(self, "Aviso", mensagem)

    def _on_aplicar_lote(self, propriedade: str, valor: Any) -> None:
        """Trata aplicação em lote."""
        selecionados = self._table_panel.obter_selecionados()

        if not selecionados:
            QMessageBox.warning(self, "Aviso", "Nenhum suporte selecionado")
            return

        resposta = QMessageBox.question(
            self,
            "Confirmação",
            f"Aplicar '{propriedade}' = {valor} em {len(selecionados)} suportes?",
            QMessageBox.Yes | QMessageBox.No
        )

        if resposta != QMessageBox.Yes:
            return

        # Executa em lote
        self._executar_edicao_lote(selecionados, propriedade, valor)

    def _executar_edicao_lote(self, suportes: List[SuporteData], propriedade: str, valor: Any) -> None:
        """Executa edição em lote em background."""
        if self._batch_worker and self._batch_worker.isRunning():
            QMessageBox.warning(self, "Aviso", "Edição em andamento")
            return

        self._progress_bar.setVisible(True)
        self._progress_bar.setRange(0, len(suportes))
        self._progress_bar.setValue(0)

        self._batch_worker = BatchEditWorker(self)
        self._batch_worker.configurar(self._repository, suportes, propriedade, valor)
        self._batch_worker.progress.connect(lambda a, t: self._progress_bar.setValue(a))
        self._batch_worker.finished.connect(self._on_lote_finalizado)
        self._batch_worker.error.connect(self._on_erro_lote)
        self._batch_worker.status.connect(self._mostrar_status)
        self._batch_worker.start()

    def _on_lote_finalizado(self, stats: Dict[str, Any]) -> None:
        """Trata finalização de edição em lote."""
        self._progress_bar.setVisible(False)

        mensagem = (
            f"Edição em lote concluída:\n"
            f"Total: {stats['total']}\n"
            f"Sucesso: {stats['sucesso']}\n"
            f"Falhas: {stats['falhas']}"
        )

        if stats['falhas'] > 0:
            detalhes = "\n\nFalhas:\n"
            for detalhe in stats['detalhes']:
                if not detalhe['sucesso']:
                    detalhes += f"- {detalhe['tag']}: {detalhe['erro']}\n"
            mensagem += detalhes

        QMessageBox.information(self, "Edição em Lote", mensagem)

    def _on_erro_lote(self, mensagem: str) -> None:
        """Trata erro na edição em lote."""
        self._progress_bar.setVisible(False)
        QMessageBox.critical(self, "Erro", mensagem)

    # === Presets e Histórico ===

    def _salvar_preset(self, nome: str, descricao: str) -> None:
        """Salva um preset."""
        if self._search_service is None:
            return

        sucesso, mensagem = self._search_service.criar_preset(nome, descricao)
        if sucesso:
            QMessageBox.information(self, "Preset", mensagem)
        else:
            QMessageBox.warning(self, "Preset", mensagem)

    def _carregar_preset(self, nome: str = "") -> None:
        """Carrega um preset."""
        if self._search_service is None:
            return

        presets = self._search_service.listar_presets()

        if not presets:
            QMessageBox.information(self, "Preset", "Nenhum preset salvo")
            return

        # Se nome não fornecido, mostra diálogo (simplificado - usar primeira)
        if not nome and presets:
            nome = presets[0]['nome']

        sucesso, mensagem, filtros = self._search_service.carregar_preset(nome)

        if sucesso:
            # Atualiza UI
            self._on_busca("", filtros)
            QMessageBox.information(self, "Preset", mensagem)
        else:
            QMessageBox.warning(self, "Preset", mensagem)

    def _gerenciar_presets(self) -> None:
        """Gerencia presets."""
        if self._search_service is None:
            return

        presets = self._search_service.listar_presets()

        if not presets:
            QMessageBox.information(self, "Presets", "Nenhum preset salvo")
            return

        # Simples - em implementação completa mostraria diálogo
        lista_texto = "\n".join([f"- {p['nome']}: {p['descricao']}" for p in presets])
        QMessageBox.information(self, "Presets Salvos", lista_texto)

    def _navegar_historico(self, direcao: int) -> None:
        """Navega no histórico."""
        if self._search_service is None:
            return

        historico = self._search_service.obter_historico()

        if not historico:
            QMessageBox.information(self, "Histórico", "Nenhum histórico disponível")
            return

        # Em implementação completa, navegaria propriamente
        QMessageBox.information(self, "Histórico", f"{len(historico)} buscas no histórico")

    # === Ordenação ===

    def _ordenar_por_tag(self) -> None:
        """Ordena a tabela por TAG."""
        self._table_panel.ordenar_por_tag()

    def _ordenar_por_tipo(self) -> None:
        """Ordena a tabela por Tipo."""
        self._table_panel.ordenar_por_tipo()

    # === Utilitários ===

    def _mostrar_status(self, mensagem: str) -> None:
        """Mostra mensagem na barra de status."""
        self._status_bar.showMessage(mensagem, 5000)

    def _mostrar_sobre(self) -> None:
        """Mostra diálogo sobre."""
        QMessageBox.about(
            self,
            "Sobre",
            "<h3>Navegação de Suportes AutoCAD</h3>"
            "<p>Aplicação para navegação e edição de suportes de tubulação.</p>"
            "<p>Versão: 1.0.0</p>"
            "<p>Desenvolvido com PySide6 e AutoCAD COM</p>"
        )

    def closeEvent(self, event) -> None:
        """Trata fechamento da janela."""
        # Cancela workers em execução
        if self._load_worker and self._load_worker.isRunning():
            self._load_worker.cancelar()
            self._load_worker.wait()

        if self._batch_worker and self._batch_worker.isRunning():
            self._batch_worker.cancelar()
            self._batch_worker.wait()

        # Desconecta do AutoCAD
        if self._repository.is_connected:
            self._repository.desconectar()

        event.accept()


def main():
    """Função principal."""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
