"""Gerenciador de presets de busca."""

import json
import os
from typing import Any, Dict, List


class PresetManager:
    """
    Gerencia presets de filtros de busca.

    Salva e carrega configurações de filtros em arquivo JSON.
    """

    def __init__(self, data_dir: str = "data"):
        """
        Inicializa o gerenciador.

        Args:
            data_dir: Diretório para salvar os presets
        """
        self._data_dir = data_dir
        self._arquivo_presets = os.path.join(data_dir, "presets.json")
        self._presets: Dict[str, Dict[str, Any]] = {}

        # Cria diretório se não existir
        os.makedirs(data_dir, exist_ok=True)

        # Carrega presets existentes
        self._carregar_arquivo()

    def _carregar_arquivo(self) -> None:
        """Carrega presets do arquivo JSON."""
        if os.path.exists(self._arquivo_presets):
            try:
                with open(self._arquivo_presets, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    self._presets = dados.get('presets', {})
            except Exception as e:
                print(f"Erro ao carregar presets: {e}")
                self._presets = {}
        else:
            self._presets = {}

    def _salvar_arquivo(self) -> bool:
        """Salva presets no arquivo JSON."""
        try:
            with open(self._arquivo_presets, 'w', encoding='utf-8') as f:
                json.dump({'presets': self._presets}, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar presets: {e}")
            return False

    def salvar(self, nome: str, dados: Dict[str, Any]) -> tuple[bool, str]:
        """
        Salva um preset.

        Args:
            nome: Nome do preset
            dados: Dicionário com dados do preset

        Returns:
            Tupla (sucesso, mensagem)
        """
        # Valida nome
        nome = nome.strip()
        if not nome:
            return False, "Nome do preset não pode ser vazio"

        # Verifica se já existe
        if nome in self._presets:
            return False, f"Preset '{nome}' já existe"

        # Salva
        self._presets[nome] = dados

        if self._salvar_arquivo():
            return True, f"Preset '{nome}' salvo com sucesso"
        else:
            del self._presets[nome]
            return False, "Erro ao salvar preset"

    def atualizar(self, nome: str, dados: Dict[str, Any]) -> tuple[bool, str]:
        """
        Atualiza um preset existente.

        Args:
            nome: Nome do preset
            dados: Novos dados

        Returns:
            Tupla (sucesso, mensagem)
        """
        if nome not in self._presets:
            return False, f"Preset '{nome}' não existe"

        self._presets[nome] = dados

        if self._salvar_arquivo():
            return True, f"Preset '{nome}' atualizado"
        else:
            return False, "Erro ao atualizar preset"

    def carregar(self, nome: str) -> tuple[bool, Any]:
        """
        Carrega um preset.

        Args:
            nome: Nome do preset

        Returns:
            Tupla (sucesso, dados_preset)
        """
        if nome not in self._presets:
            return False, f"Preset '{nome}' não encontrado"

        return True, self._presets[nome]

    def deletar(self, nome: str) -> tuple[bool, str]:
        """
        Deleta um preset.

        Args:
            nome: Nome do preset

        Returns:
            Tupla (sucesso, mensagem)
        """
        if nome not in self._presets:
            return False, f"Preset '{nome}' não encontrado"

        del self._presets[nome]

        if self._salvar_arquivo():
            return True, f"Preset '{nome}' deletado"
        else:
            self._presets[nome] = {}
            return False, "Erro ao deletar preset"

    def listar_todos(self) -> List[Dict[str, Any]]:
        """
        Lista todos os presets.

        Returns:
            Lista de presets com informações resumidas
        """
        resultado = []

        for nome, dados in self._presets.items():
            resultado.append({
                'nome': nome,
                'descricao': dados.get('descricao', ''),
                'data_criacao': dados.get('data_criacao', ''),
                'quantidade_filtros': len(dados.get('filtros', []))
            })

        return sorted(resultado, key=lambda x: x['nome'])

    def existe(self, nome: str) -> bool:
        """
        Verifica se um preset existe.

        Args:
            nome: Nome do preset

        Returns:
            True se existe
        """
        return nome in self._presets

    def renomear(self, nome_antigo: str, nome_novo: str) -> tuple[bool, str]:
        """
        Renomeia um preset.

        Args:
            nome_antigo: Nome atual
            nome_novo: Novo nome

        Returns:
            Tupla (sucesso, mensagem)
        """
        if nome_antigo not in self._presets:
            return False, f"Preset '{nome_antigo}' não encontrado"

        if nome_novo in self._presets:
            return False, f"Já existe um preset com o nome '{nome_novo}'"

        self._presets[nome_novo] = self._presets.pop(nome_antigo)

        if self._salvar_arquivo():
            return True, f"Preset renomeado para '{nome_novo}'"
        else:
            self._presets[nome_antigo] = self._presets.pop(nome_novo)
            return False, "Erro ao renomear preset"
