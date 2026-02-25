"""Gerenciador de histórico de buscas."""

import json
import os
from datetime import datetime
from typing import Any, Dict, List, Tuple


class HistoryManager:
    """
    Gerencia o histórico de buscas realizadas.

    Mantém um histórico circular com tamanho máximo.
    """

    MAX_ENTRADAS = 100  # Máximo de entradas no histórico

    def __init__(self, data_dir: str = "data"):
        """
        Inicializa o gerenciador.

        Args:
            data_dir: Diretório para salvar o histórico
        """
        self._data_dir = data_dir
        self._arquivo_historico = os.path.join(data_dir, "history.json")
        self._historico: List[Dict[str, Any]] = []
        self._indice_atual = -1  # Para navegação anterior/próximo

        # Cria diretório se não existir
        os.makedirs(data_dir, exist_ok=True)

        # Carrega histórico existente
        self._carregar_arquivo()

    def _carregar_arquivo(self) -> None:
        """Carrega histórico do arquivo JSON."""
        if os.path.exists(self._arquivo_historico):
            try:
                with open(self._arquivo_historico, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    self._historico = dados.get('historico', [])
            except Exception as e:
                print(f"Erro ao carregar histórico: {e}")
                self._historico = []
        else:
            self._historico = []

    def _salvar_arquivo(self) -> bool:
        """Salva histórico no arquivo JSON."""
        try:
            with open(self._arquivo_historico, 'w', encoding='utf-8') as f:
                json.dump({'historico': self._historico}, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar histórico: {e}")
            return False

    def adicionar(self, dados_busca: Dict[str, Any]) -> bool:
        """
        Adiciona uma entrada ao histórico.

        Args:
            dados_busca: Dicionário com dados da busca

        Returns:
            True se salvo com sucesso
        """
        # Adiciona timestamp se não existir
        if 'data_hora' not in dados_busca:
            dados_busca['data_hora'] = datetime.now().isoformat()

        # Adiciona ao início da lista
        self._historico.insert(0, dados_busca)

        # Mantém apenas o máximo
        if len(self._historico) > self.MAX_ENTRADAS:
            self._historico = self._historico[:self.MAX_ENTRADAS]

        # Atualiza índice
        self._indice_atual = -1

        return self._salvar_arquivo()

    def listar(self, limite: int = 50) -> List[Dict[str, Any]]:
        """
        Lista entradas do histórico.

        Args:
            limite: Número máximo de entradas

        Returns:
            Lista de buscas do histórico
        """
        return self._historico[:limite]

    def limpar(self) -> tuple[bool, str]:
        """
        Limpa todo o histórico.

        Returns:
            Tupla (sucesso, mensagem)
        """
        self._historico.clear()
        self._indice_atual = -1

        if self._salvar_arquivo():
            return True, "Histórico limpo"
        else:
            return False, "Erro ao limpar histórico"

    def obter_anterior(self) -> Dict[str, Any]:
        """
        Retorna a entrada anterior do histórico (navegação).

        Returns:
            Entrada anterior ou dict vazio
        """
        if not self._historico:
            return {}

        self._indice_atual += 1

        if self._indice_atual >= len(self._historico):
            self._indice_atual = len(self._historico) - 1

        return self._historico[self._indice_atual]

    def obter_proximo(self) -> Dict[str, Any]:
        """
        Retorna a próxima entrada do histórico (navegação).

        Returns:
            Próxima entrada ou dict vazio
        """
        if not self._historico or self._indice_atual <= 0:
            self._indice_atual = -1
            return {}

        self._indice_atual -= 1

        if self._indice_atual < 0:
            self._indice_atual = -1
            return {}

        return self._historico[self._indice_atual]

    def resetar_navegacao(self) -> None:
        """Reseta o índice de navegação."""
        self._indice_atual = -1

    def obter_por_indice(self, indice: int) -> Dict[str, Any]:
        """
        Obtém entrada por índice.

        Args:
            indice: Índice da entrada

        Returns:
            Entrada ou dict vazio
        """
        if 0 <= indice < len(self._historico):
            return self._historico[indice]
        return {}

    def remover(self, indice: int) -> tuple[bool, str]:
        """
        Remove uma entrada do histórico.

        Args:
            indice: Índice da entrada

        Returns:
            Tupla (sucesso, mensagem)
        """
        if 0 <= indice < len(self._historico):
            removida = self._historico.pop(indice)

            # Ajusta índice se necessário
            if self._indice_atual >= len(self._historico):
                self._indice_atual = len(self._historico) - 1

            if self._salvar_arquivo():
                return True, "Entrada removida"
            else:
                self._historico.insert(indice, removida)
                return False, "Erro ao remover entrada"

        return False, "Índice inválido"

    def buscar(self, texto: str) -> List[Dict[str, Any]]:
        """
        Busca entradas no histórico.

        Args:
            texto: Texto para buscar

        Returns:
            Lista de entradas que contêm o texto
        """
        texto_lower = texto.lower()
        resultado = []

        for entrada in self._historico:
            # Busca em texto geral
            if texto_lower in entrada.get('texto_geral', '').lower():
                resultado.append(entrada)
                continue

            # Busca em filtros
            for filtro in entrada.get('filtros', []):
                if (texto_lower in filtro.get('campo', '').lower() or
                    texto_lower in str(filtro.get('valor', '')).lower()):
                    resultado.append(entrada)
                    break

        return resultado

    def exportar(self, caminho: str) -> tuple[bool, str]:
        """
        Exporta o histórico para um arquivo.

        Args:
            caminho: Caminho do arquivo

        Returns:
            Tupla (sucesso, mensagem)
        """
        try:
            with open(caminho, 'w', encoding='utf-8') as f:
                json.dump({'historico': self._historico}, f, indent=2, ensure_ascii=False)
            return True, f"Histórico exportado para {caminho}"
        except Exception as e:
            return False, f"Erro ao exportar: {e}"

    def importar(self, caminho: str, substituir: bool = False) -> tuple[bool, str]:
        """
        Importa histórico de um arquivo.

        Args:
            caminho: Caminho do arquivo
            substituir: Se True, substitui o histórico atual

        Returns:
            Tupla (sucesso, mensagem)
        """
        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                dados = json.load(f)

            novo_historico = dados.get('historico', [])

            if substituir:
                self._historico = novo_historico
            else:
                # Adiciona ao início
                self._historico = novo_historico + self._historico

            # Limita tamanho
            if len(self._historico) > self.MAX_ENTRADAS:
                self._historico = self._historico[:self.MAX_ENTRADAS]

            self._indice_atual = -1

            if self._salvar_arquivo():
                return True, f"Histórico importado: {len(novo_historico)} entradas"
            else:
                return False, "Erro ao salvar histórico importado"

        except Exception as e:
            return False, f"Erro ao importar: {e}"

    @property
    def tamanho(self) -> int:
        """Retorna o número de entradas no histórico."""
        return len(self._historico)
