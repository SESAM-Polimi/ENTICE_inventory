"""Inspect the canonical registry and read-only inventory diagnostics."""
import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'src'))
from entice_inventory.core.registry import Registry, CHECKS, inspect_inventory


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--registry-dir',type=Path)
    parser.add_argument('--output',type=Path,help='Write complete JSON diagnostics; refuses an existing path.')
    sub = parser.add_subparsers(dest='command',required=True)
    check = sub.add_parser('check')
    check.add_argument('--checks',nargs='*',choices=CHECKS,default=list(CHECKS))
    check.add_argument('--target',default='GTAP12')
    sector = sub.add_parser('sector')
    sector.add_argument('value')
    sector.add_argument('--namespace',default='sector_id')
    sector.add_argument('--scope')
    sector.add_argument('--target',default='GTAP12')
    region = sub.add_parser('region')
    region.add_argument('value')
    region.add_argument('--geography',default='GTAP12')
    inventory = sub.add_parser('inspect-inventory')
    inventory.add_argument('path',type=Path)
    inventory.add_argument('--checks',nargs='*',choices=['metadata','coverage'],default=['metadata','coverage'])
    inventory.add_argument('--geography',default='GTAP12')
    args = parser.parse_args()
    try:
        registry = Registry.load(args.registry_dir)
        if args.command=='check':
            result = registry.check(args.checks,args.target)
        elif args.command=='sector':
            record = registry.sector(args.value,args.namespace)
            result = {'sector':record,'mapping':registry.parent(record['id'],args.target,args.scope)}
        elif args.command=='region':
            result = {'geography':args.geography,'identifier':args.value,'members':registry.members(args.value,args.geography)}
        else:
            result = inspect_inventory(args.path,registry,args.checks,args.geography)
        if args.output:
            args.output.parent.mkdir(parents=True,exist_ok=True)
            with args.output.open('x') as stream:
                json.dump(result,stream,indent=2,ensure_ascii=False,allow_nan=False)
                stream.write('\n')
            summary = {k:v for k,v in result.items() if k in {'status','executed_checks','skipped_checks','warning_count','warnings_by_check','coverage'}}
            summary['details'] = str(args.output)
            print(json.dumps(summary,indent=2))
        else:
            print(json.dumps(result,indent=2,ensure_ascii=False,allow_nan=False))
    except Exception as exc:
        print(json.dumps({'status':'failed','error':str(exc)}),file=sys.stderr)
        return 1
    return 0


if __name__=='__main__':
    raise SystemExit(main())
