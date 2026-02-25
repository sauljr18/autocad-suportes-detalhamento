"""Módulo services - Camada de serviço da aplicação."""

from .search_service import SearchService
from .preset_manager import PresetManager
from .history_manager import HistoryManager

__all__ = ['SearchService', 'PresetManager', 'HistoryManager']
