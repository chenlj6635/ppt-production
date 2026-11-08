import hashlib
import json
from pathlib import Path
import sys
import tempfile
import zipfile


def load_backend():
    engines = Path(__file__).resolve().parents[1] / 'assets' / 'engines'
    manifest = json.loads((engines / 'manifest.json').read_text(encoding='utf-8'))
    archive = engines / manifest['archive']
    if hashlib.sha256(archive.read_bytes()).hexdigest() != manifest['sha256']:
        raise ValueError('Bundled native engine archive failed its SHA-256 check')
    cache = Path(tempfile.gettempdir()) / 'ppt-production-native' / manifest['sha256'][:20]
    cache.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive) as bundle:
        if set(bundle.namelist()) != set(manifest['files']):
            raise ValueError('Bundled native engine file inventory differs from its manifest')
        for name, digest in manifest['files'].items():
            target = (cache / name).resolve()
            if not target.is_relative_to(cache.resolve()):
                raise ValueError('Invalid archive path')
            if target.is_file() and hashlib.sha256(target.read_bytes()).hexdigest() == digest:
                continue
            data = bundle.read(name)
            if hashlib.sha256(data).hexdigest() != digest:
                raise ValueError('Native engine member failed its SHA-256 check: ' + name)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    scripts = cache / 'scripts'
    if str(scripts) not in sys.path:
        sys.path.insert(0, str(scripts))
    from attribution_guard import require_skill_integrity
    require_skill_integrity()
    return manifest
