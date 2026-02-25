"""Handler para erros COM do AutoCAD."""

import time
import pythoncom
from typing import Any, Callable, List, Optional


class COMErrorInfo:
    """Informações sobre um erro COM."""

    def __init__(self, attempt: int, operation: str, exception: Exception):
        self.attempt = attempt
        self.operation = operation
        self.exception = exception
        self.timestamp = time.time()

    def __str__(self) -> str:
        return f"Tentativa {self.attempt} - {self.operation}: {type(self.exception).__name__}: {self.exception}"


class COMErrorHandler:
    """
    Handler para erros COM RPC_E_CALL_REJECTED do AutoCAD.

    O AutoCAD pode rejeitar chamadas COM quando está ocupado. Este handler
    implementa retry automático com delays crescentes.
    """

    # HRESULT para RPC_E_CALL_REJECTED (AutoCAD ocupado)
    RPC_E_CALL_REJECTED = -2147418111
    # HRESULT para RPC_E_SERVERCALL_RETRYLATER
    RPC_E_SERVERCALL_RETRYLATER = -2147418110

    def __init__(self, max_retries: int = 3, base_delay: float = 0.5):
        """
        Inicializa o handler.

        Args:
            max_retries: Número máximo de tentativas
            base_delay: Delay base em segundos
        """
        self.max_retries = max_retries
        self.base_delay = base_delay
        self._error_log: List[COMErrorInfo] = []

    @property
    def error_log(self) -> List[COMErrorInfo]:
        """Histórico de erros."""
        return self._error_log.copy()

    def clear_log(self) -> None:
        """Limpa o histórico de erros."""
        self._error_log.clear()

    def execute_with_retry(
        self,
        func: Callable[[], Any],
        operation_name: str = "Operation",
        max_retries: Optional[int] = None
    ) -> Any:
        """
        Executa uma função com retry para erros RPC.

        Args:
            func: Função a executar
            operation_name: Nome da operação para log
            max_retries: Override do número máximo de tentativas

        Returns:
            Resultado da função

        Raises:
            Exception: Se todas as tentativas falharem
        """
        retries = max_retries if max_retries is not None else self.max_retries

        for attempt in range(1, retries + 1):
            try:
                return func()
            except pythoncom.com_error as e:
                error_info = COMErrorInfo(attempt, operation_name, e)
                self._error_log.append(error_info)

                # Verifica se é um erro recuperável
                if e.hresult in (self.RPC_E_CALL_REJECTED, self.RPC_E_SERVERCALL_RETRYLATER):
                    if attempt < retries:
                        # Delay crescente: 0.5s, 1s, 1.5s, ...
                        delay = self.base_delay * attempt
                        time.sleep(delay)
                        continue

                # Se não for recuperável ou última tentativa, propaga
                raise

            except Exception as e:
                error_info = COMErrorInfo(attempt, operation_name, e)
                self._error_log.append(error_info)
                raise

        raise Exception(f"{operation_name} falhou após {retries} tentativas")

    @staticmethod
    def execute_static(func: Callable[[], Any], operation_name: str = "Operation", max_retries: int = 3) -> Any:
        """
        Método estático para executar com retry (compatibilidade com código legado).

        Args:
            func: Função a executar
            operation_name: Nome da operação para log
            max_retries: Número máximo de tentativas

        Returns:
            Resultado da função

        Raises:
            Exception: Se todas as tentativas falharem
        """
        handler = COMErrorHandler(max_retries=max_retries)
        return handler.execute_with_retry(func, operation_name)

    def is_recoverable_com_error(self, exception: Exception) -> bool:
        """
        Verifica se a exceção é um erro COM recuperável.

        Args:
            exception: Exceção a verificar

        Returns:
            True se for um erro recuperável
        """
        if isinstance(exception, pythoncom.com_error):
            return exception.hresult in (self.RPC_E_CALL_REJECTED, self.RPC_E_SERVERCALL_RETRYLATER)
        return False

    def get_retry_suggestion(self) -> str:
        """
        Retorna sugestão baseada nos erros registrados.

        Returns:
            String com sugestão de ação
        """
        if not self._error_log:
            return "Nenhum erro registrado."

        recoverable_count = sum(
            1 for e in self._error_log
            if self.is_recoverable_com_error(e.exception)
        )

        total = len(self._error_log)

        if recoverable_count == total:
            return (
                f"Todos os {total} erros foram recuperáveis via retry. "
                "O AutoCAD pode estar sob carga pesada. Considere aumentar delays."
            )
        elif recoverable_count > 0:
            return (
                f"{recoverable_count} de {total} erros foram recuperáveis. "
                f"{total - recoverable_count} erros não foram recuperáveis."
            )
        else:
            return "Nenhum erro recuperável detectado. Verifique a lógica da aplicação."


# Singleton global para uso em toda aplicação
_global_error_handler: Optional[COMErrorHandler] = None


def get_global_error_handler() -> COMErrorHandler:
    """Retorna o handler global de erros COM."""
    global _global_error_handler
    if _global_error_handler is None:
        _global_error_handler = COMErrorHandler()
    return _global_error_handler


def execute_with_retry(func: Callable[[], Any], operation_name: str = "Operation", max_retries: int = 3) -> Any:
    """
    Função de conveniência para executar com retry usando o handler global.

    Args:
        func: Função a executar
        operation_name: Nome da operação para log
        max_retries: Número máximo de tentativas

    Returns:
        Resultado da função
    """
    return get_global_error_handler().execute_with_retry(func, operation_name, max_retries)
