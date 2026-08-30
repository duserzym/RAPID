"""Device ownership guards for shared RapidPy hardware services."""
from __future__ import annotations

from dataclasses import dataclass
import threading


class DeviceOwnershipError(RuntimeError):
    """Raised when a device is already reserved by another owner."""


@dataclass(frozen=True)
class DeviceLease:
    """Context-style handle proving ownership of one logical device."""

    resource: str
    owner: str
    _manager: "DeviceOwnershipManager"

    def release(self) -> None:
        """Release the owned device back to the manager."""
        self._manager.release(self.resource, self.owner)


class DeviceOwnershipManager:
    """Tracks exclusive ownership of logical devices.

    A device is identified by a stable string key and can only be actively owned by
    one owner at a time.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        # resource -> [owner, reentrant_count]
        self._owners: dict[str, tuple[str, int]] = {}

    def is_owned(self, resource: str) -> bool:
        """Return True when *resource* is currently owned."""
        with self._lock:
            return resource in self._owners

    def owner_of(self, resource: str) -> str | None:
        """Return the active owner id, or None."""
        with self._lock:
            entry = self._owners.get(resource)
            return entry[0] if entry else None

    def acquire(self, resource: str, owner: str, *, allow_reentrant: bool = True) -> DeviceLease:
        """Claim a device for the given owner.

        Parameters
        ----------
        resource:
            Stable logical device name (e.g. ``"squid"`` or ``"af_demag"``).
        owner:
            Caller identity used for conflict reporting and reentrant checks.
        allow_reentrant:
            When True, the same owner can reacquire a lock it already holds.
        """
        resource = str(resource)
        owner = str(owner)

        with self._lock:
            current_entry = self._owners.get(resource)
            current = current_entry[0] if current_entry else None
            if current is None:
                self._owners[resource] = (owner, 1)
                return DeviceLease(resource, owner, self)
            if current == owner and allow_reentrant:
                _owner, count = self._owners[resource]
                self._owners[resource] = (_owner, count + 1)
                return DeviceLease(resource, owner, self)
            raise DeviceOwnershipError(
                f"Resource '{resource}' is in use by '{current}', "
                f"not '{owner}'"
            )

    def release(self, resource: str, owner: str) -> None:
        """Release a claimed device; invalid callers are ignored."""
        resource = str(resource)
        owner = str(owner)

        with self._lock:
            current_entry = self._owners.get(resource)
            if current_entry is None or current_entry[0] != owner:
                return

            owner_name, count = current_entry
            if count <= 1:
                del self._owners[resource]
            else:
                self._owners[resource] = (owner_name, count - 1)
