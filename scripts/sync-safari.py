#!/usr/bin/env python3
"""Copy shared extension sources into the macOS and iPad Xcode projects."""
import json
import plistlib
import re
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROJECT = ROOT / 'safari/NYCU E3 Helper'
FILES = ['background.js', 'content.js', 'i18n.js', 'desktop-notifications.js', 'notification-api.js', 'notification-engine.js',
         'notification-settings.js', 'notification-settings.html',
         'notification-settings.css', 'jszip.min.js', '128.png', '_locales']

def sync_project(project, extension, native_notifications):
    resources = project / extension / 'Resources'
    for name in FILES:
        source, target = ROOT / name, resources / name
        if source.is_dir():
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(source, target)
        else:
            shutil.copy2(source, target)
    manifest = json.loads((ROOT / 'manifest.json').read_text())
    manifest['permissions'] = [p for p in manifest['permissions'] if p not in ('downloads', 'notifications')]
    if native_notifications:
        manifest['permissions'].append('nativeMessaging')
    manifest['options_ui'].pop('open_in_tab', None)
    # The AI providers use hosted APIs; no local Ollama service is needed.
    manifest['host_permissions'] = [p for p in manifest['host_permissions'] if 'localhost' not in p]
    (resources / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
    pbx = next(project.glob('*.xcodeproj')) / 'project.pbxproj'
    source = pbx.read_text()
    if source.startswith('<?xml'):
        data = plistlib.loads(pbx.read_bytes())
        for obj in data['objects'].values():
            settings = obj.get('buildSettings', {})
            if 'MARKETING_VERSION' in settings:
                settings['MARKETING_VERSION'] = manifest['version']
        pbx.write_bytes(plistlib.dumps(data, sort_keys=False))
    else:
        source = re.sub(r'MARKETING_VERSION = [^;]+;', f"MARKETING_VERSION = {manifest['version']};", source)
        pbx.write_text(source)
    print(f"{project.name} resources synced: {manifest['version']}")

def sync():
    sync_project(PROJECT, 'NYCU E3 Helper Extension', True)
    ios = ROOT / 'safari/ios/NYCU E3 Helper iOS'
    if ios.exists():
        sync_project(ios, 'NYCU E3 Helper iOS Extension', False)

if __name__ == '__main__':
    sync()
