#!/usr/bin/env python3
"""Record benchmark context and report advisory changes against a compatible run."""
import argparse
import csv
import hashlib
import html
import io
import json
import os
import re
from pathlib import Path
import subprocess
import zipfile

ENV_FIELDS = ('schema_version', 'sample_interval_ms', 'java_version', 'os', 'arch',
              'max_heap_bytes', 'processors', 'gc_collectors')
NUMBERS = ('rows', 'elapsed_nanos', 'output_bytes', 'heap_baseline_bytes', 'heap_peak_bytes',
           'temp_baseline_bytes', 'temp_peak_bytes', 'temp_remaining_bytes', 'samples',
           'sample_interval_ms', 'max_heap_bytes', 'processors')
METRICS = {'elapsed_nanos': 'Time', 'heap_peak_bytes': 'Heap peak', 'temp_peak_bytes': 'Temp peak'}
WORKLOAD_FILES = ('BenchmarkTest.java', 'BenchmarkProbe.java', 'BenchmarkReport.java')


def record_context(root):
    digest = hashlib.sha256()
    for name in WORKLOAD_FILES:
        source = root / 'kit/src/test/java/io/github/dornol/excelkit/benchmark' / name
        digest.update(name.encode() + b'\0' + source.read_bytes() + b'\0')
    cpu = ''
    if Path('/proc/cpuinfo').exists():
        cpu = next((line.split(':', 1)[1].strip() for line in Path('/proc/cpuinfo').read_text().splitlines()
                    if line.startswith('model name')), '')
    return {'schema_version': 1, 'workload_sha256': digest.hexdigest(),
            'runner': {key: os.environ.get(key, '') for key in
                       ('RUNNER_OS', 'RUNNER_ARCH', 'RUNNER_ENVIRONMENT', 'ImageOS', 'ImageVersion')},
            'cpu_model': cpu}


def load_csv(text):
    reader = csv.DictReader(io.StringIO(text))
    required = set(ENV_FIELDS + NUMBERS + ('scenario', 'revision'))
    if not reader.fieldnames or not required.issubset(reader.fieldnames):
        raise ValueError('CSV is missing required benchmark columns')
    if len(set(reader.fieldnames)) != len(reader.fieldnames):
        raise ValueError('CSV has duplicate columns')
    rows = {}
    for raw in reader:
        if None in raw or any(value is None for value in raw.values()):
            raise ValueError('CSV row has an invalid number of columns')
        if raw['schema_version'] != '1' or not raw['scenario'] or raw['scenario'] in rows:
            raise ValueError('Unsupported schema or duplicate/empty scenario')
        row = dict(raw)
        for key in NUMBERS:
            row[key] = int(raw[key])
            if row[key] < 0:
                raise ValueError(f'Negative benchmark value: {key}')
        if any(row[key] == 0 for key in ('rows', 'elapsed_nanos', 'sample_interval_ms', 'processors', 'max_heap_bytes')):
            raise ValueError('Workload, elapsed time and environment limits must be positive')
        if row['samples'] < 2 or row['heap_peak_bytes'] < row['heap_baseline_bytes']:
            raise ValueError('Invalid sample count or heap peak')
        if row['temp_peak_bytes'] < max(row['temp_baseline_bytes'], row['temp_remaining_bytes']):
            raise ValueError('Invalid temporary-file peak')
        rows[row['scenario']] = row
    if not rows:
        raise ValueError('CSV has no benchmark results')
    first = next(iter(rows.values()))
    if any(any(row[key] != first[key] for key in ENV_FIELDS + ('revision',)) for row in rows.values()):
        raise ValueError('CSV mixes environments or revisions')
    return rows


def load_context(text):
    context = json.loads(text)
    if (not isinstance(context, dict) or context.get('schema_version') != 1
            or not isinstance(context.get('workload_sha256'), str)
            or not re.fullmatch(r'[0-9a-f]{64}', context['workload_sha256'])
            or not isinstance(context.get('runner'), dict)
            or not all(isinstance(key, str) and isinstance(value, str) for key, value in context['runner'].items())
            or not isinstance(context.get('cpu_model'), str)):
        raise ValueError('Invalid or unsupported benchmark context')
    return context


def mismatch(current, context, baseline, baseline_context):
    if context.get('schema_version') != 1 or baseline_context.get('schema_version') != 1:
        return 'context schema differs'
    if not context.get('workload_sha256') or context['workload_sha256'] != baseline_context.get('workload_sha256'):
        return 'workload or measurement implementation differs'
    for key in ('runner', 'cpu_model'):
        if context.get(key) != baseline_context.get(key):
            return f'{key} differs'
    if list(current) != list(baseline):
        return 'scenario set or execution order differs'
    for name, row in current.items():
        previous = baseline[name]
        if row['rows'] != previous['rows']:
            return f'row count differs for {name}'
        for key in ENV_FIELDS:
            if row[key] != previous[key]:
                return f'{key} differs'
    return None


def changes(current, baseline):
    results = []
    for name, row in current.items():
        metrics = {}
        for key in METRICS:
            old, new = baseline[name][key], row[key]
            metrics[key] = {'before': old, 'after': new,
                            'percent': (new - old) * 100 / old if old else None,
                            'increased': new > old}
        results.append({'scenario': name, 'metrics': metrics})
    return results


def gh_api(path, binary=False):
    result = subprocess.run(['gh', 'api', path], capture_output=True, check=False)
    if result.returncode:
        # Do not echo API error bodies or token-bearing redirect URLs.
        raise RuntimeError(f'GitHub API request failed (exit {result.returncode})')
    return result.stdout if binary else json.loads(result.stdout)


def artifact_reports(archive):
    with zipfile.ZipFile(io.BytesIO(archive)) as bundle:
        reports = {}
        for name in ('resources.csv', 'context.json'):
            matches = [item for item in bundle.infolist() if Path(item.filename).name == name]
            if len(matches) != 1 or matches[0].file_size > 2_000_000:
                raise ValueError(f'Artifact must contain exactly one bounded {name}')
            reports[name] = bundle.read(matches[0]).decode('utf-8')
        return load_csv(reports['resources.csv']), load_context(reports['context.json'])


def select_baseline(current, context, repository, run_id, api=gh_api):
    run = api(f'repos/{repository}/actions/runs/{run_id}')
    runs = api(f'repos/{repository}/actions/workflows/benchmark.yml/runs?branch=main&event=push&status=success&per_page=30')
    skipped = []
    candidates = sorted(runs['workflow_runs'], key=lambda item: item['created_at'], reverse=True)
    for candidate in candidates:
        if (str(candidate['id']) == str(run_id) or candidate['created_at'] >= run['created_at']
                or candidate['event'] != 'push' or candidate['head_branch'] != 'main'
                or candidate['conclusion'] != 'success'
                or (candidate.get('head_repository') or {}).get('full_name') != repository):
            continue
        candidate_id = candidate['id']
        try:
            artifacts = api(f'repos/{repository}/actions/runs/{candidate_id}/artifacts?per_page=100')['artifacts']
            artifact = next((item for item in artifacts if item['name'] == 'benchmark-results' and not item['expired']), None)
            if artifact is None:
                skipped.append(f'Run {candidate_id}: report artifact missing or expired')
                continue
            previous, previous_context = artifact_reports(api(f'repos/{repository}/actions/artifacts/{artifact["id"]}/zip', binary=True))
            reason = mismatch(current, context, previous, previous_context)
            if reason:
                skipped.append(f'Run {candidate_id}: {reason}')
                continue
            # Artifact provenance must match the successful main run we selected.
            if any(row['revision'] != candidate['head_sha'] for row in previous.values()):
                skipped.append(f'Run {candidate_id}: artifact revision differs from workflow revision')
                continue
            return previous, f'https://github.com/{repository}/actions/runs/{candidate_id}', skipped
        except (RuntimeError, ValueError, KeyError, zipfile.BadZipFile, UnicodeError) as error:
            skipped.append(f'Run {candidate_id}: {error}')
    return None, None, skipped


def escape(value):
    return html.escape(str(value)).replace('|', '&#124;').replace('\n', ' ').replace('\r', ' ')


def metric_text(metric, key):
    def value(number):
        return f'{number / 1_000_000_000:.3f}s' if key == 'elapsed_nanos' else f'{number / 1048576:.2f} MiB'
    percent = metric['percent']
    delta = f'{percent:+.2f}%' if percent is not None else ('new usage' if metric['after'] else 'unchanged')
    return f'{value(metric["before"])} → {value(metric["after"])} ({delta})'


def summarize(current, baseline, baseline_url, notes):
    environment = next(iter(current.values()))
    lines = ['## Benchmark observations', '',
             'Advisory only. Sampled peaks are lower bounds; shared-runner noise and GC can affect differences.', '',
             f'Environment: Java {escape(environment["java_version"])}; {escape(environment["os"])}/{escape(environment["arch"])}; '
             f'{environment["processors"]} processors; {environment["max_heap_bytes"] / 1048576:.0f} MiB heap limit; '
             f'{environment["sample_interval_ms"]} ms sampling.', '']
    comparisons = changes(current, baseline) if baseline is not None else []
    if baseline is not None:
        lines += [f'Baseline: [{escape(baseline_url)}]({escape(baseline_url)})' if baseline_url else 'Baseline: local report', '',
                  'No calibrated regression threshold is applied. Upward changes are observations, not confirmed regressions.', '',
                  '| Scenario | Time | Heap peak | Temp peak |', '| --- | --- | --- | --- |']
        for item in comparisons:
            values = [metric_text(item['metrics'][key], key) for key in METRICS]
            lines.append('| ' + ' | '.join([escape(item['scenario'])] + values) + ' |')
    else:
        lines += ['No compatible previous successful main baseline was available; current results are retained for future comparison.', '',
                  '| Scenario | Time | Heap peak | Temp peak |', '| --- | --- | --- | --- |']
        for name, row in current.items():
            lines.append(f'| {escape(name)} | {row["elapsed_nanos"] / 1e9:.3f}s | {row["heap_peak_bytes"] / 1048576:.2f} MiB | {row["temp_peak_bytes"] / 1048576:.2f} MiB |')
    increases = sum(any(metric['increased'] for metric in row['metrics'].values()) for row in comparisons)
    lines += ['', f'Scenarios with observed increases: {increases}' if baseline is not None else 'Comparison status: baseline unavailable', '']
    if notes:
        lines += ['<details><summary>Baseline selection details</summary>', ''] + ['- ' + escape(note) for note in notes] + ['', '</details>', '']
    return '\n'.join(lines), comparisons, increases


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    context_parser = sub.add_parser('context')
    context_parser.add_argument('--output', type=Path, required=True)
    compare = sub.add_parser('compare')
    compare.add_argument('--current', type=Path, required=True)
    compare.add_argument('--context', type=Path, required=True)
    compare.add_argument('--output-dir', type=Path, required=True)
    compare.add_argument('--github', action='store_true')
    compare.add_argument('--baseline', type=Path)
    compare.add_argument('--baseline-context', type=Path)
    args = parser.parse_args()
    if args.command == 'context':
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(record_context(Path(__file__).resolve().parent.parent), indent=2) + '\n')
        return
    current = load_csv(args.current.read_text())
    context = load_context(args.context.read_text())
    baseline, url, notes = None, None, []
    if args.baseline:
        if not args.baseline_context:
            parser.error('--baseline requires --baseline-context')
        candidate = load_csv(args.baseline.read_text())
        reason = mismatch(current, context, candidate, load_context(args.baseline_context.read_text()))
        if reason:
            notes.append(reason)
        else:
            baseline = candidate
    elif args.github:
        try:
            baseline, url, notes = select_baseline(current, context, os.environ['GITHUB_REPOSITORY'], os.environ['GITHUB_RUN_ID'])
        except (RuntimeError, KeyError, ValueError, OSError) as error:
            notes.append(f'Baseline lookup unavailable: {error}')
            print('::warning::Benchmark baseline lookup unavailable; see comparison summary.')
    summary, comparisons, increases = summarize(current, baseline, url, notes)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    (args.output_dir / 'comparison.md').write_text(summary)
    (args.output_dir / 'comparison.json').write_text(json.dumps({'baseline_url': url, 'baseline_available': baseline is not None,
        'notes': notes, 'scenarios_with_increases': increases, 'comparisons': comparisons}, indent=2) + '\n')
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a') as output:
            output.write(summary + '\n')
    if increases:
        print(f'::warning::Observed upward benchmark changes in {increases} scenarios; no calibrated threshold is applied. See the comparison summary.')
    print(f'Benchmark comparison saved to {args.output_dir / "comparison.md"}')


if __name__ == '__main__':
    main()
