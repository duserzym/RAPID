"""Enter native backend ownership on the worker that performs all stage I/O."""
from contextlib import nullcontext
from inspect import getattr_static


def worker_claim(backend):
    # Only use a declared capability, never one invented by dynamic proxies.
    declared = getattr_static(backend, 'queue_worker_claim', None)
    claim = getattr(backend, 'queue_worker_claim', None) if declared is not None else None
    return claim() if callable(claim) else nullcontext()
