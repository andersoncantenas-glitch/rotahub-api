# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Optional


@dataclass
class AppContext:
    """Container leve para configuração e hooks de runtime.

    - config: objeto de configuração (ex.: APP_CONFIG)
    - hooks: funções de runtime não modularizadas
    """

    config: Optional[Any] = None
    hooks: Dict[str, Callable[..., Any]] = field(default_factory=dict)
