"""独立版的本地数据目录；不读取其他应用的配置。"""
import json
import os
import tempfile
from pathlib import Path

DATA = Path(os.environ.get('PAPERCC_DATA_DIR', '~/.papercc')).expanduser().resolve()
CACHE = DATA / 'cache'


def _desktop_base():
    return DATA


def _read_json(path, default):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return default


def _write_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    fd, name = tempfile.mkstemp(dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as handle:
            json.dump(obj, handle, ensure_ascii=False, indent=2)
        os.replace(name, path)
    finally:
        Path(name).unlink(missing_ok=True)


def load_config():
    return _read_json(DATA / 'config.json', {})
