"""Tests for config.AppConfig — persistent JSON configuration.

Covers:
  - First launch (no file) returns defaults
  - get / set round-trip
  - Favorite font add / remove / is_favorite / duplicate prevention
  - Corrupt JSON file recovery
  - Missing file graceful fallback
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from config import AppConfig


def _make_config(tmp_path: Path) -> AppConfig:
    return AppConfig(config_dir=tmp_path)


# -- basic get/set ---------------------------------------------------------

def test_get_missing_key_returns_default(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    assert cfg.get("nonexistent") is None
    assert cfg.get("nonexistent", 42) == 42


def test_set_and_get(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.set("key1", "value1")
    assert cfg.get("key1") == "value1"


def test_persists_across_instances(tmp_path: Path) -> None:
    cfg1 = _make_config(tmp_path)
    cfg1.set("fruit", "apple")

    cfg2 = _make_config(tmp_path)
    assert cfg2.get("fruit") == "apple"


# -- favorite fonts --------------------------------------------------------

def test_favorite_fonts_empty_by_default(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    assert cfg.get_favorite_fonts() == []


def test_add_favorite_font(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.add_favorite_font("Arial")
    cfg.add_favorite_font("Georgia")
    assert cfg.get_favorite_fonts() == ["Arial", "Georgia"]


def test_add_duplicate_font_ignored(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.add_favorite_font("Arial")
    cfg.add_favorite_font("Arial")
    assert cfg.get_favorite_fonts() == ["Arial"]


def test_remove_favorite_font(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.add_favorite_font("Arial")
    cfg.add_favorite_font("Georgia")
    cfg.remove_favorite_font("Arial")
    assert cfg.get_favorite_fonts() == ["Georgia"]


def test_remove_nonexistent_font_no_error(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.remove_favorite_font("NoSuchFont")  # should not raise
    assert cfg.get_favorite_fonts() == []


def test_is_favorite_font(tmp_path: Path) -> None:
    cfg = _make_config(tmp_path)
    cfg.add_favorite_font("Calibri")
    assert cfg.is_favorite_font("Calibri") is True
    assert cfg.is_favorite_font("Arial") is False


# -- error recovery --------------------------------------------------------

def test_corrupt_json_returns_defaults(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text("{invalid json", encoding="utf-8")

    cfg = _make_config(tmp_path)
    assert cfg.get("key") is None
    assert cfg.get_favorite_fonts() == []


def test_non_dict_json_returns_defaults(tmp_path: Path) -> None:
    config_file = tmp_path / "config.json"
    config_file.write_text('"just a string"', encoding="utf-8")

    cfg = _make_config(tmp_path)
    assert cfg.get("key") is None


def test_missing_file_returns_defaults(tmp_path: Path) -> None:
    cfg = AppConfig(config_dir=tmp_path / "nonexistent_dir")
    assert cfg.get("key") is None
    assert cfg.get_favorite_fonts() == []
