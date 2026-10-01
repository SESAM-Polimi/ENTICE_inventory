"""Versioned activity identities, explicit geography and NACE-led target adapters.

Legacy labels and published parents are evidence, never fallback classifications.
Ambiguous aliases fail explicitly. Optional diagnostic findings do not block reads.
"""
from __future__ import annotations

from collections import Counter
import hashlib
from importlib.metadata import distribution, PackageNotFoundError
import json
from pathlib import Path
import re
import sys

from entice_inventory.core.paths import DATA_DIR


CHECKS = ('classification', 'identity', 'regions')


def _default_directory():
    checkout = DATA_DIR/'registry'
    if checkout.is_dir():
        return checkout
    # pip --target places data files beside the package; normal installations
    # place them below the environment's data prefix.
    for prefix in (Path(__file__).resolve().parents[2], Path(sys.prefix)):
        candidate = prefix/'share/entice-inventory/registry'
        if candidate.is_dir():
            return candidate
    # Wheels install the same source files once under share/. Resolve the actual
    # distribution location rather than assuming a particular virtualenv prefix.
    try:
        installed = distribution('entice-inventory')
        for path in installed.files or []:
            if str(path).endswith('share/entice-inventory/registry/sectors.json'):
                candidate = Path(installed.locate_file(path)).parent
                if candidate.is_dir():
                    return candidate
    except PackageNotFoundError:
        pass
    raise FileNotFoundError('Registry data are unavailable; supply an explicit registry directory')


def normalise(value):
    """Ignore case and whitespace only; punctuation can distinguish activities."""
    return ' '.join(str(value).split()).casefold()


def finding(check_id, message, action, *, sector=None, region=None,
            source_record=None, observed=None, reference=None, severity='warning'):
    return dict(check_id=check_id, severity=severity, sector=sector, region=region,
                source_record=source_record, observed=observed, reference=reference,
                tolerance=None, message=message, suggested_action=action)


def _index(rows, key, label):
    result = {}
    for row in rows:
        value = row[key]
        if not isinstance(value, str) or not value.strip() or value in result:
            raise ValueError(f'Duplicate or empty {label}: {value!r}')
        result[value] = row
    return result


class Registry:
    def __init__(self, sectors, regions, adapters):
        for document in [sectors, regions, *adapters]:
            if document.get('schema_version') != 1:
                raise ValueError('Unsupported registry schema_version')
        self.sector_document = sectors
        self.region_document = regions
        self.sectors = _index(sectors['sectors'], 'id', 'sector identity')
        self.geographies = _index(regions['geographies'], 'id', 'geography')
        self.adapters = _index(adapters, 'target', 'target adapter')
        self.aliases = sectors['aliases']
        for alias in self.aliases:
            if alias['sector_id'] not in self.sectors or not alias['namespace'] or not alias['value']:
                raise ValueError(f'Invalid alias: {alias}')
        for sector in self.sectors.values():
            if not re.fullmatch(r'[A-Z][A-Z0-9_]*', sector['id']):
                raise ValueError(f'Invalid canonical sector identifier: {sector["id"]}')
            if sector['lifecycle'] not in {'published_d24', 'catalogue_only_not_activated'}:
                raise ValueError(f'Invalid lifecycle: {sector["lifecycle"]}')
            c = sector['classification']
            if c['system'] != 'NACE' or c['version'] != 'Rev. 2':
                raise ValueError(f'Unsupported classification for {sector["id"]}')
            if c['status'] not in {'unreviewed', 'reviewed_scope', 'scope_required'}:
                raise ValueError(f'Invalid classification status for {sector["id"]}')
            scopes = _index(c['scopes'], 'id', 'activity scope')
            if c['status'] == 'reviewed_scope' and len(scopes) != 1:
                raise ValueError('A reviewed activity must have exactly one scope')
            if c['status'] == 'unreviewed' and scopes:
                raise ValueError('Unreviewed classification cannot carry resolved scopes')
            if scopes and not c['evidence']:
                raise ValueError('A classification scope requires evidence')
            for scope in scopes.values():
                if not scope['codes'] or any(not re.fullmatch(r'\d{2}(?:\.\d{2})?', n) for n in scope['codes']):
                    raise ValueError(f'Invalid NACE codes: {scope["codes"]}')
            for relation in sector['related_records']:
                if relation['sector_id'] not in self.sectors:
                    raise ValueError(f'Unknown related sector: {relation}')
        for g in self.geographies.values():
            members = _index(g['regions'], 'id', 'region')
            groups = _index(g['groups'], 'id', 'region group')
            if set(members) & set(groups):
                raise ValueError('Region and group identifiers must be distinct')
            for group in groups.values():
                if len(group['members']) != len(set(group['members'])) or set(group['members']) - set(members):
                    raise ValueError(f'Invalid members in {g["id"]}/{group["id"]}')
            if 'GLOBAL' not in groups or set(groups['GLOBAL']['members']) != set(members):
                raise ValueError(f'{g["id"]}/GLOBAL must contain every target region exactly once')
        for adapter in self.adapters.values():
            if adapter['geography'] not in self.geographies:
                raise ValueError('Adapter refers to an unknown geography')
            bridge = _index(adapter['bridge'], 'nace_code', 'NACE bridge rule')
            parent_catalogue = _index(adapter['parent_catalogue'], 'id', 'target parent')
            records = _index(adapter['records'], 'sector_id', 'target sector record')
            if set(records) != set(self.sectors):
                raise ValueError('Target records must cover every registered sector identity')
            if any(not r['parent'] or not r['evidence'] for r in bridge.values()):
                raise ValueError('A bridge rule requires a parent and source evidence')
            if any(r['parent'] not in parent_catalogue for r in bridge.values()):
                raise ValueError('A bridge rule refers to an unknown target parent')
        self.fingerprints = {}

    @classmethod
    def load(cls, directory=None):
        directory = Path(directory) if directory is not None else _default_directory()
        paths = [directory/'sectors.json', directory/'regions.json']
        paths.extend(sorted(p for p in directory.glob('*.json') if p not in paths))
        documents = [json.loads(p.read_text()) for p in paths]
        registry = cls(documents[0], documents[1], documents[2:])
        registry.fingerprints = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
        return registry

    def sector(self, value, namespace='sector_id'):
        if namespace == 'sector_id':
            candidates = {code for code in self.sectors if normalise(code) == normalise(value)}
        else:
            candidates = {a['sector_id'] for a in self.aliases
                          if a['namespace'] == namespace and normalise(a['value']) == normalise(value)}
        if len(candidates) != 1:
            raise LookupError(f'{namespace}:{value!r} resolves to {sorted(candidates)}; use an unambiguous sector_id')
        return self.sectors[next(iter(candidates))]

    def parent(self, sector_id, target='GTAP12', scope=None):
        sector = self.sector(sector_id)
        adapter = self.adapters[target]
        classification = sector['classification']
        record = next(r for r in adapter['records'] if r['sector_id'] == sector['id'])
        result = dict(sector_id=sector['id'], target=target, parent=None,
                      status='unreviewed', scope=None, nace_codes=[],
                      published_parent=record['published_parent'], requires_rebuild=False,
                      available_scopes=classification['scopes'])
        if classification['version'] != adapter['nace_version']:
            return {**result, 'status':'classification_version_mismatch'}
        scopes = {s['id']:s for s in classification['scopes']}
        if scope is not None and scope not in scopes:
            raise LookupError(f'Unknown activity scope {sector_id}/{scope}')
        if not scopes:
            return result
        if scope is None and classification['status'] == 'scope_required':
            return {**result, 'status':'scope_required'}
        selected = scopes[scope] if scope is not None else next(iter(scopes.values()))
        bridge = {r['nace_code']:r['parent'] for r in adapter['bridge']}
        result.update(scope=selected['id'], nace_codes=selected['codes'])
        if any(n not in bridge for n in selected['codes']):
            return {**result, 'status':'bridge_missing'}
        parents = {bridge[n] for n in selected['codes']}
        if len(parents) != 1:
            return {**result, 'status':'multiple_parents'}
        parent = next(iter(parents))
        return {**result, 'parent':parent,
                'status':'conditional_on_scope' if classification['status']=='scope_required' else 'resolved',
                'requires_rebuild': record['published_parent'] is not None and parent != record['published_parent']}

    def regions(self, geography='GTAP12'):
        return self.geographies[geography]['regions']

    def members(self, value, geography='GTAP12'):
        g = self.geographies[geography]
        direct = [r['id'] for r in g['regions'] if normalise(value) in {normalise(r['id']), normalise(r['name'])}]
        groups = [r for r in g['groups'] if normalise(r['id']) == normalise(value)]
        if len(direct) + len(groups) != 1:
            raise LookupError(f'Unknown or ambiguous region/group: {geography}/{value}')
        return direct if direct else list(groups[0]['members'])

    def check(self, checks=CHECKS, target='GTAP12'):
        checks = set(checks)
        if checks - set(CHECKS):
            raise ValueError(f'Unknown checks: {checks - set(CHECKS)}')
        adapter = self.adapters[target]
        issues = []
        if 'classification' in checks:
            for sector in self.sectors.values():
                resolution = self.parent(sector['id'], target)
                if resolution['parent'] is None:
                    issues.append(finding('classification.'+resolution['status'],
                        f'No automatic {target} parent is available for {sector["id"]}.',
                        'Resolve the activity scope and versioned NACE bridge using source evidence.',
                        sector=sector['id'], source_record=sector['source_records'],
                        observed=sector['classification'], reference=resolution))
                elif resolution['requires_rebuild']:
                    issues.append(finding('classification.parent_changed',
                        f'{sector["id"]}: published {resolution["published_parent"]}; NACE-derived {resolution["parent"]}.',
                        'Rebuild parent-dependent coefficients and outputs before exporting this mapping.',
                        sector=sector['id'], observed=resolution['published_parent'], reference=resolution['parent']))
                if sector['classification'].get('review_status') in {'scope_name_correction', 'parent_and_source_correction', 'duplicate_file_review'}:
                    issues.append(finding('classification.source_alignment',
                        sector['classification']['note'],
                        'Align the name, cost profile, production and trade scope before rebuilding.',
                        sector=sector['id'],source_record=sector['source_records']))
        if 'identity' in checks:
            grouped = {}
            for a in self.aliases:
                grouped.setdefault((a['namespace'],normalise(a['value'])),set()).add(a['sector_id'])
            for (namespace,value),codes in grouped.items():
                if len(codes)>1:
                    issues.append(finding('identity.ambiguous_alias',
                        f'{namespace}:{value} refers to multiple activities.',
                        'Use the stable sector identity and preserve source context.', observed=sorted(codes)))
            for sector in self.sectors.values():
                if sector['lifecycle'] != 'published_d24':
                    issues.append(finding('identity.not_activated',
                        f'{sector["id"]} is catalogued but has no published D2.4 inventory.',
                        'Keep inactive until its inclusion or withdrawal is established.', sector=sector['id']))
                if sector['related_records']:
                    issues.append(finding('identity.unmerged_records',
                        f'{sector["id"]} has related records that are not interchangeable.',
                        'Compare activity scope and numerical sources before merging identities.',
                        sector=sector['id'], observed=sector['related_records']))
        if 'regions' in checks:
            # Membership integrity is checked on load; this records diagnostic execution.
            self.members('GLOBAL', adapter['geography'])
        return self.result(issues, checks, CHECKS, {'target':target})

    def result(self, issues, checks, available, extra=None):
        return {'status':'completed_with_warnings' if issues else 'completed',
                'executed_checks':sorted(checks), 'skipped_checks':sorted(set(available)-set(checks)),
                'registry_sha256':self.fingerprints, 'warning_count':len(issues),
                'warnings_by_check':dict(Counter(i['check_id'] for i in issues)),
                'warnings':issues, **(extra or {})}


def inspect_inventory(path, registry=None, checks=('metadata','coverage'), geography='GTAP12'):
    """Read a MARIO workbook without changing it or filling missing observations."""
    from openpyxl import load_workbook
    registry = registry or Registry.load()
    checks = set(checks)
    if checks - {'metadata','coverage'}:
        raise ValueError('Inventory checks must be metadata and/or coverage')
    path = Path(path)
    workbook = load_workbook(path, read_only=True, data_only=True)
    issues = []
    coverage = {}
    try:
        master = list(workbook['Master'].values)
        headers = list(master[0])
        si, ri = headers.index('Sector'), headers.index('Region')
        codes = sorted({str(r[si]).strip() for r in master[1:] if r[si]})
        sector_id = codes[0] if len(codes) == 1 else None
        if 'metadata' in checks:
            for code in codes:
                try:
                    registry.sector(code)
                except LookupError as exc:
                    issues.append(finding('metadata.unknown_sector',str(exc),'Register the identity before rebuilding.',sector=code,source_record='Master!Sector'))
                    continue
                if 'Parent Sector' in headers:
                    expected = registry.parent(code)['parent']
                    parents = sorted({str(r[headers.index('Parent Sector')]).strip() for r in master[1:] if r[si]==code})
                    if expected is not None and parents != [expected]:
                        issues.append(finding('metadata.parent_mismatch',
                            'The inventory parent differs from its NACE-derived parent.',
                            'Rebuild all parent-dependent quantities; changing this label alone is insufficient.',
                            sector=code,source_record='Master!Parent Sector',observed=parents,reference=expected))
            if 'Summary' in workbook and sector_id in registry.sectors:
                summary = list(workbook['Summary'].values)
                for i,row in enumerate(summary,1):
                    if row[0] == 'Sector' and len(row)>2:
                        expected = registry.sectors[sector_id]['name']
                        if str(row[1]).strip()!=sector_id or normalise(row[2])!=normalise(expected):
                            issues.append(finding('metadata.sector_description',
                                'Summary code or description differs from the registered activity.',
                                'Reconcile the source metadata with its activity scope.',sector=sector_id,
                                source_record=f'Summary!B{i}:C{i}',observed=list(row[1:3]),reference=[sector_id,expected]))
        if 'coverage' in checks:
            known = set(registry.members('GLOBAL',geography))
            clusters = {}
            if 'Regions Clusters' in workbook:
                rows = list(workbook['Regions Clusters'].values)
                for col,name in enumerate(rows[0]):
                    if not name:
                        continue
                    if str(name) in clusters:
                        raise ValueError(f'Duplicate cluster header: {name}')
                    values = [str(r[col]).strip() for r in rows[1:] if r[col] is not None]
                    clusters[str(name)] = set(values)
                    if len(values)!=len(set(values)) or set(values)-known:
                        issues.append(finding('coverage.invalid_cluster',
                            f'{name} contains duplicate or unknown region codes.',
                            'Correct the membership against the declared target geography.',
                            sector=sector_id,source_record=f'Regions Clusters!{name}',observed=values))
            global_members = clusters.get('GLOBAL',set())
            missing = sorted(known-global_members)
            if missing:
                issues.append(finding('coverage.global_missing_regions',
                    f'GLOBAL omits {len(missing)} target regions: {", ".join(missing)}.',
                    'Rebuild the source inventory using the canonical region registry; do not assume missing observations are zero.',
                    sector=sector_id,source_record='Regions Clusters!GLOBAL',observed=sorted(global_members),reference=sorted(known)))
            producing = set()
            for row in master[1:]:
                value = str(row[ri]).strip() if row[ri] is not None else ''
                if value in known:
                    producing.add(value)
                elif value in clusters:
                    producing.update(clusters[value] & known)
                elif value:
                    issues.append(finding('coverage.unknown_producing_region',
                        f'Master region {value} is neither a region nor a defined source cluster.',
                        'Define explicit source membership.',sector=sector_id,region=value,source_record='Master!Region'))
            if known-producing:
                issues.append(finding('coverage.producing_regions_missing',
                    'The Master rows do not cover every target region.',
                    'Review intended coverage; record estimates separately from observations.',
                    sector=sector_id,source_record='Master!Region',observed=sorted(known-producing)))
            coverage = {'geography':geography,'target_region_count':len(known),
                        'global_region_count':len(global_members & known),
                        'global_missing_regions':missing,'producing_missing_regions':sorted(known-producing)}
    finally:
        workbook.close()
    return registry.result(issues, checks, ('metadata','coverage'),
        {'source':str(path),'source_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'coverage':coverage})
