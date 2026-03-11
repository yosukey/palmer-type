"""Persistent JSON-backed application configuration.

Stores settings in a platform-appropriate user data directory so that both
installer and portable (one-file) builds share the same config.

    Windows : %APPDATA%\\palmer-type\\config.json
    macOS   : ~/Library/Application Support/palmer-type/config.json
    Linux   : ~/.config/palmer-type/config.json
"""

from __future__ import annotations

import json
import logging
import os
import platform
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

_APP_NAME = "palmer-type"


class AppConfig:
    """Lazily loads and persists a flat JSON config dict."""

    def __init__(self, config_dir: Path | None = None) -> None:
        """*config_dir* overrides the default platform directory (for tests)."""
        self._config_dir = config_dir
        self._data: dict[str, Any] | None = None  # lazy

    # -- path resolution ---------------------------------------------------

    @staticmethod
    def _default_config_dir() -> Path:
        system = platform.system()
        if system == "Windows":
            base = os.environ.get("APPDATA")
            if base:
                return Path(base) / _APP_NAME
            return Path.home() / "AppData" / "Roaming" / _APP_NAME
        if system == "Darwin":
            return Path.home() / "Library" / "Application Support" / _APP_NAME
        # Linux / others
        xdg = os.environ.get("XDG_CONFIG_HOME")
        if xdg:
            return Path(xdg) / _APP_NAME
        return Path.home() / ".config" / _APP_NAME

    def _config_path(self) -> Path:
        d = self._config_dir if self._config_dir is not None else self._default_config_dir()
        return d / "config.json"

    # -- load / save --------------------------------------------------------

    def _ensure_loaded(self) -> dict[str, Any]:
        if self._data is None:
            self._data = self._load()
        return self._data

    def _load(self) -> dict[str, Any]:
        path = self._config_path()
        if not path.exists():
            return {}
        try:
            text = path.read_text(encoding="utf-8")
            data = json.loads(text)
            if isinstance(data, dict):
                return data
            logger.warning("Config is not a JSON object, resetting: %s", path)
            return {}
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("Failed to load config, using defaults: %s", exc)
            return {}

    def _save(self, data: dict[str, Any]) -> None:
        path = self._config_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, path)
        except OSError as exc:
            logger.warning("Failed to save config: %s", exc)

    # -- generic get/set ----------------------------------------------------

    def get(self, key: str, default: Any = None) -> Any:
        return self._ensure_loaded().get(key, default)

    def set(self, key: str, value: Any) -> None:
        data = self._ensure_loaded()
        data[key] = value
        self._save(data)

    # -- favorite fonts helpers ---------------------------------------------

    _FAV_KEY = "favorite_fonts"

    def get_favorite_fonts(self) -> list[str]:
        val = self.get(self._FAV_KEY, [])
        if isinstance(val, list):
            return list(val)
        return []

    def add_favorite_font(self, name: str) -> None:
        fonts = self.get_favorite_fonts()
        if name not in fonts:
            fonts.append(name)
            self.set(self._FAV_KEY, fonts)

    def remove_favorite_font(self, name: str) -> None:
        fonts = self.get_favorite_fonts()
        if name in fonts:
            fonts.remove(name)
            self.set(self._FAV_KEY, fonts)

    def is_favorite_font(self, name: str) -> bool:
        return name in set(self.get_favorite_fonts())

    def clear_favorite_fonts(self) -> None:
        """Remove all favorite fonts from the config."""
        data = self._ensure_loaded()
        data.pop(self._FAV_KEY, None)
        self._save(data)

    # -- reset helpers ------------------------------------------------------

    def delete(self) -> None:
        """Delete the config file and reset in-memory state to empty."""
        self._data = {}
        path = self._config_path()
        try:
            path.unlink(missing_ok=True)
        except OSError as exc:
            logger.warning("Failed to delete config file: %s", exc)
