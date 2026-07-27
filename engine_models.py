"""Shared engine states and the conversion backend contract."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import AbstractSet, Callable, Protocol, Union, runtime_checkable


class EngineState(str, Enum):
    """Portable result states for shallow and deep engine probes."""

    CHECKING = "checking"
    AVAILABLE = "available"
    MISSING = "missing"
    UNVERIFIED = "unverified"
    PERMISSION_DENIED = "permission_denied"
    LAUNCH_FAILED = "launch_failed"
    TIMEOUT = "timeout"
    UNSUPPORTED = "unsupported"


_INSTALLED_STATES = frozenset(
    {
        EngineState.AVAILABLE,
        EngineState.UNVERIFIED,
        EngineState.PERMISSION_DENIED,
        EngineState.LAUNCH_FAILED,
        EngineState.TIMEOUT,
    }
)


@dataclass(frozen=True)
class EngineStatus:
    """Result of probing one conversion engine.

    ``installed`` means the engine was discovered on disk or could be
    launched far enough to fail for another reason. ``usable`` is deliberately
    stricter and is true only after an available result. A probe is
    ``complete`` once it is no longer in the transient checking state;
    ``unverified`` is therefore a valid completed shallow probe.
    """

    state: EngineState
    detail: str = ""

    def __post_init__(self) -> None:
        if not isinstance(self.state, EngineState):
            object.__setattr__(self, "state", EngineState(self.state))
        if self.detail is None:
            object.__setattr__(self, "detail", "")

    @property
    def installed(self) -> bool:
        return self.state in _INSTALLED_STATES

    @property
    def usable(self) -> bool:
        return self.state is EngineState.AVAILABLE

    @property
    def complete(self) -> bool:
        return self.state is not EngineState.CHECKING


PathLike = Union[str, Path]
ProgressCallback = Callable[[str, int], None]


@runtime_checkable
class ConversionBackend(Protocol):
    """Contract implemented by platform-native and built-in converters."""

    @property
    def id(self) -> str:
        ...

    @property
    def display_name(self) -> str:
        ...

    @property
    def directions(self) -> AbstractSet[str]:
        ...

    def probe(self, deep: bool = False) -> EngineStatus:
        ...

    def convert(
        self,
        source: PathLike,
        output: PathLike,
        progress: ProgressCallback,
    ) -> None:
        ...
