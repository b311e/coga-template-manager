#!/usr/bin/env python3
"""
Update foundation/partials/partials-manifest.json to reflect snippet files present in foundation/partials/.

Usage:
  python3 src/scripts/update_partials_manifest.py [--dry-run]

The script will back up the existing manifest to partials-manifest.json.bak before writing.
"""
import os
import json
import xml.etree.ElementTree as ET
import argparse


def local_name(tag):
    if tag is None:
        return None
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def infer_target_and_tags(root):
    # search children for common markers
    tags = set()
    target = None
    for elem in root.iter():
        ln = local_name(elem.tag)
        if ln in ('numbering', 'abstractNum', 'num'):
            target = 'word:numbering'
            tags.add('numbering')
        if ln == 'style':
            # styles part
            if target is None:
                target = 'word:styles'
            tags.add('styles')
    return target, sorted(tags)


def parse_snippet_file(path, repo_root):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
    except Exception:
        return None

    ln = local_name(root.tag)
    rel = os.path.relpath(path, repo_root).replace('\\', '/')

    if ln == 'snippet':
        sid = root.attrib.get('id') or os.path.splitext(os.path.basename(path))[0]
        target, tags = infer_target_and_tags(root)
        return {
            'id': sid,
            'path': rel,
            'target': target or '',
            'tags': tags or [],
            'status': 'active'
        }

    # not a snippet root; try to infer from content
    sid = os.path.splitext(os.path.basename(path))[0]
    target, tags = infer_target_and_tags(root)
    return {
        'id': sid,
        'path': rel,
        'target': target or '',
        'tags': tags or [],
        'status': 'active'
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true', help='Print the manifest JSON but do not write')
    args = parser.parse_args()

    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    partials_dir = os.path.join(repo_root, 'foundation', 'partials')
    manifest_path = os.path.join(partials_dir, 'partials-manifest.json')

    if not os.path.isdir(partials_dir):
        print('partials directory not found:', partials_dir)
        return 2

    entries = []
    for dirpath, dirnames, filenames in os.walk(partials_dir):
        for fn in sorted(filenames):
            if not fn.lower().endswith('.xml') and not fn.lower().endswith('.xmlx'):
                continue
            full = os.path.join(dirpath, fn)
            e = parse_snippet_file(full, repo_root)
            if e:
                entries.append(e)

    # sort entries by id then path
    entries.sort(key=lambda x: (x.get('id') or '', x.get('path') or ''))

    manifest = {'snippets': entries}

    out = json.dumps(manifest, indent=2)

    if args.dry_run:
        print(out)
        return 0

    # backup
    if os.path.exists(manifest_path):
        bak = manifest_path + '.bak'
        try:
            os.replace(manifest_path, bak)
            print('Backup saved to', bak)
        except Exception as ex:
            print('Warning: could not create backup:', ex)

    # write new manifest
    try:
        with open(manifest_path, 'w', encoding='utf-8') as f:
            f.write(out + '\n')
        print('Wrote', manifest_path)
    except Exception as ex:
        print('Error writing manifest:', ex)
        return 3

    print('Found', len(entries), 'snippet(s)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
