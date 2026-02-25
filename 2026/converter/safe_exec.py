# -*- coding: utf-8 -*-
"""Safe execution helper extracted from office_converter.py."""


def safe_exec(
    func,
    *args,
    retries=3,
    is_running_getter=None,
    sleep_fn=None,
    randint_fn=None,
    com_error_cls=None,
    rpc_server_busy_code=None,
    **kwargs,
):
    if is_running_getter is None:
        is_running_getter = lambda: True
    if sleep_fn is None:
        import time as _time

        sleep_fn = _time.sleep
    if randint_fn is None:
        import random as _random

        randint_fn = _random.randint

    retryable_exceptions = (OSError, RuntimeError, TypeError, ValueError)
    if (
        isinstance(com_error_cls, type)
        and issubclass(com_error_cls, BaseException)
        and com_error_cls not in retryable_exceptions
    ):
        retryable_exceptions = retryable_exceptions + (com_error_cls,)

    for attempt in range(retries + 1):
        if not is_running_getter():
            raise RuntimeError("program stopped")
        try:
            return func(*args, **kwargs)
        except retryable_exceptions as e:
            if com_error_cls is not None and isinstance(e, com_error_cls):
                error_code = getattr(e, "hresult", None)
                if rpc_server_busy_code is not None and error_code == rpc_server_busy_code:
                    sleep_fn(randint_fn(2, 5))
                    continue
                if attempt < retries:
                    sleep_fn(1)
                    continue
                raise RuntimeError(f"COM error ({error_code}): {e}")
            if attempt < retries:
                sleep_fn(1)
                continue
            raise
