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


class DeviceLeaseGroup:
    """One atomic reservation; release never decrements a later reservation."""
    def __init__(self, resources, owner, manager):
        self.resources, self.owner, self._manager = tuple(resources), owner, manager
        self._released = False

    def release(self):
        with self._manager._lock:
            if self._released:
                return
            self._released = True
            for resource in self.resources:
                self._manager._release_owned(resource, self.owner)


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
            self._release_owned(resource, owner)

    def _release_owned(self, resource, owner):
        current_entry = self._owners.get(resource)
        if current_entry is None or current_entry[0] != owner:
            return
        owner_name, count = current_entry
        if count <= 1:
            del self._owners[resource]
        else:
            self._owners[resource] = (owner_name, count - 1)

    def acquire_many(self, resources, owner, *, allow_reentrant=True):
        resources, owner = tuple(str(resource) for resource in resources), str(owner)
        if not resources or len(set(resources)) != len(resources):
            raise DeviceOwnershipError('A device reservation requires distinct resources.')
        with self._lock:
            for resource in resources:
                current = self._owners.get(resource)
                if current is not None and (current[0] != owner or not allow_reentrant):
                    raise DeviceOwnershipError(f"Resource '{resource}' is in use by '{current[0]}', not '{owner}'")
            for resource in resources:
                current = self._owners.get(resource)
                self._owners[resource] = (owner, current[1] + 1 if current else 1)
        return DeviceLeaseGroup(resources, owner, self)
