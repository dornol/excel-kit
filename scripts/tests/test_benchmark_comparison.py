import copy
import importlib.util
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

SPEC = importlib.util.spec_from_file_location('benchmark_comparison', Path(__file__).parents[1] / 'benchmark_comparison.py')
bench = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(bench)
CSV = '''schema_version,scenario,rows,elapsed_nanos,output_bytes,heap_baseline_bytes,heap_peak_bytes,temp_baseline_bytes,temp_peak_bytes,temp_remaining_bytes,samples,sample_interval_ms,java_version,os,arch,max_heap_bytes,processors,gc_collectors,revision
1,write,1000,1000000000,2048,100,1000,0,0,0,100,10,21.0.1,Linux,amd64,1073741824,4,G1,old
'''
CONTEXT = {'schema_version': 1, 'workload_sha256': 'a' * 64, 'runner': {'ImageVersion': 'image1'}, 'cpu_model': 'cpu1'}


def archive(csv=CSV, context=CONTEXT):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as bundle:
        bundle.writestr('benchmark/resources.csv', csv)
        if context is not None:
            bundle.writestr('benchmark/context.json', json.dumps(context))
    return output.getvalue()


def candidate(run_id, created, **overrides):
    result = dict(id=run_id, created_at=created, event='push', head_branch='main',
                  conclusion='success', head_repository={'full_name': 'owner/repo'}, head_sha='old')
    result.update(overrides)
    return result


class ComparisonTest(unittest.TestCase):
    def setUp(self):
        self.baseline = bench.load_csv(CSV)
        self.current = copy.deepcopy(self.baseline)

    def test_percentages_and_zero_baseline_are_defined(self):
        self.current['write']['elapsed_nanos'] = 1200000000
        self.current['write']['heap_peak_bytes'] = 900
        self.current['write']['temp_peak_bytes'] = 1024
        result = bench.changes(self.current, self.baseline)[0]['metrics']
        self.assertEqual(20, result['elapsed_nanos']['percent'])
        self.assertEqual(-10, result['heap_peak_bytes']['percent'])
        self.assertIsNone(result['temp_peak_bytes']['percent'])
        self.assertTrue(result['temp_peak_bytes']['increased'])
        text, _, increases = bench.summarize(self.current, self.baseline, None, [])
        self.assertIn('+20.00%', text)
        self.assertIn('new usage', text)
        self.assertEqual(1, increases)

    def test_identical_results_have_no_increase(self):
        _, _, increases = bench.summarize(self.current, self.baseline, None, [])
        self.assertEqual(0, increases)

    def test_compatibility_checks_every_environment_field(self):
        for field in bench.ENV_FIELDS:
            with self.subTest(field=field):
                changed = copy.deepcopy(self.baseline)
                changed['write'][field] = 'different'
                self.assertIn(field, bench.mismatch(self.current, CONTEXT, changed, CONTEXT))
        for field in ('workload_sha256', 'runner', 'cpu_model'):
            with self.subTest(field=field):
                changed = copy.deepcopy(CONTEXT)
                changed[field] = 'different'
                self.assertIsNotNone(bench.mismatch(self.current, CONTEXT, self.baseline, changed))

    def test_changed_row_counts_and_scenario_sets_are_not_compared(self):
        self.baseline['write']['rows'] = 2000
        self.assertIn('row count', bench.mismatch(self.current, CONTEXT, self.baseline, CONTEXT))
        self.assertIn('scenario', bench.mismatch(self.current, CONTEXT, {}, CONTEXT))

    def test_no_baseline_is_an_explicit_initialization_state(self):
        text, results, increases = bench.summarize(self.current, None, None, ['no history'])
        self.assertIn('No compatible', text)
        self.assertIn('no history', text)
        self.assertEqual([], results)
        self.assertEqual(0, increases)

    def test_invalid_reports_do_not_silently_become_baselines(self):
        for invalid in (CSV + CSV.splitlines()[1] + '\n', CSV.replace('1000000000', '-1'),
                        CSV.replace('1000000000', 'nan'), CSV.replace(',100,1000,', ',2000,1000,'),
                        CSV.splitlines()[0] + '\n', CSV.replace('1,write,', '9,write,'),
                        CSV.replace(',old\n', '\n')):
            with self.subTest(invalid=invalid):
                self.assertRaises(ValueError, bench.load_csv, invalid)

    def test_html_and_annotation_like_scenario_text_is_escaped(self):
        row = self.current.pop('write')
        self.current['<b>|\n::warning::injected'] = row
        text, _, _ = bench.summarize(self.current, None, None, [])
        self.assertNotIn('<b>', text)
        self.assertIn('&#124;', text)
        self.assertNotIn('\n::warning::', text)

    def test_duplicate_artifact_filenames_are_rejected_without_extracting(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w') as bundle:
            bundle.writestr('one/resources.csv', CSV)
            bundle.writestr('../../resources.csv', CSV)
            bundle.writestr('context.json', json.dumps(CONTEXT))
        self.assertRaises(ValueError, bench.artifact_reports, output.getvalue())

    def test_workload_hash_tracks_benchmark_sources(self):
        context = bench.record_context(Path(__file__).parents[2])
        self.assertEqual(64, len(context['workload_sha256']))
        self.assertEqual(1, context['schema_version'])

    def test_malformed_context_is_rejected(self):
        for invalid in ([], {}, {**CONTEXT, 'workload_sha256': 'not-a-hash'}, {**CONTEXT, 'runner': []}):
            self.assertRaises(ValueError, bench.load_context, json.dumps(invalid))

    def test_cli_warns_but_succeeds_and_saves_machine_readable_results(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / 'current.csv').write_text(CSV.replace('1000000000', '1200000000'))
            (root / 'baseline.csv').write_text(CSV)
            (root / 'context.json').write_text(json.dumps(CONTEXT))
            summary = root / 'step-summary.md'
            process = subprocess.run([sys.executable, SPEC.origin, 'compare', '--current', str(root / 'current.csv'),
                '--context', str(root / 'context.json'), '--baseline', str(root / 'baseline.csv'),
                '--baseline-context', str(root / 'context.json'), '--output-dir', str(root / 'report')],
                env={**os.environ, 'GITHUB_STEP_SUMMARY': str(summary)}, capture_output=True, text=True)
            self.assertEqual(0, process.returncode, process.stderr)
            self.assertIn('::warning::', process.stdout)
            self.assertIn('+20.00%', summary.read_text())
            result = json.loads((root / 'report/comparison.json').read_text())
            self.assertTrue(result['baseline_available'])
            self.assertEqual(1, result['scenarios_with_increases'])

    def test_cli_rejects_stale_current_results_before_api_lookup(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / 'current.csv').write_text(CSV)
            (root / 'context.json').write_text(json.dumps(CONTEXT))
            process = subprocess.run([sys.executable, SPEC.origin, 'compare', '--current', str(root / 'current.csv'),
                '--context', str(root / 'context.json'), '--output-dir', str(root / 'report'), '--github'],
                env={**os.environ, 'GITHUB_SHA': 'new'}, capture_output=True, text=True)
            self.assertNotEqual(0, process.returncode)
            self.assertIn('Current benchmark revision differs', process.stderr)


class BaselineSelectionTest(unittest.TestCase):
    def setUp(self):
        self.current = bench.load_csv(CSV)
        self.calls = []
        self.runs = []
        self.reports = {}
        self.expired = set()

    def api(self, path, binary=False):
        self.calls.append(path)
        if path.endswith('/runs/100'):
            return {'created_at': '2026-11-02T02:00:00Z'}
        if '/workflows/' in path:
            return {'workflow_runs': self.runs}
        if path.endswith('/artifacts?per_page=100'):
            run_id = int(path.split('/runs/')[1].split('/')[0])
            return {'artifacts': [{'id': run_id, 'name': 'benchmark-results', 'expired': run_id in self.expired}]}
        if path.endswith('/zip'):
            return self.reports[int(path.split('/artifacts/')[1].split('/')[0])]
        raise AssertionError(path)

    def test_chooses_latest_compatible_successful_main_run(self):
        self.runs = [candidate(10, '2026-11-01T01:00:00Z'), candidate(11, '2026-11-01T02:00:00Z')]
        self.reports = {10: archive(), 11: archive()}
        _, url, notes = bench.select_baseline(self.current, CONTEXT, 'owner/repo', '100', self.api)
        self.assertTrue(url.endswith('/11'))
        self.assertFalse(notes)

    def test_skips_expired_missing_context_and_incompatible_reports(self):
        self.runs = [candidate(i, f'2026-11-01T0{i}:00:00Z') for i in range(1, 5)]
        self.expired.add(4)
        self.reports = {3: archive(context=None), 2: archive(context={**CONTEXT, 'cpu_model': 'other'}), 1: archive()}
        _, url, notes = bench.select_baseline(self.current, CONTEXT, 'owner/repo', '100', self.api)
        self.assertTrue(url.endswith('/1'))
        self.assertEqual(3, len(notes))

    def test_rejects_pr_fork_failed_current_and_future_runs(self):
        self.runs = [candidate(1, '2026-11-01T01:00:00Z', event='pull_request'),
                     candidate(2, '2026-11-01T01:00:00Z', head_repository={'full_name': 'fork/repo'}),
                     candidate(3, '2026-11-01T01:00:00Z', conclusion='failure'),
                     candidate(100, '2026-11-02T01:00:00Z'),
                     candidate(101, '2026-11-03T01:00:00Z')]
        previous, url, _ = bench.select_baseline(self.current, CONTEXT, 'owner/repo', '100', self.api)
        self.assertIsNone(previous)
        self.assertIsNone(url)
        self.assertFalse(any('/artifacts' in call for call in self.calls))

    def test_revision_must_match_the_selected_main_run(self):
        self.runs = [candidate(1, '2026-11-01T01:00:00Z', head_sha='different')]
        self.reports = {1: archive()}
        previous, _, notes = bench.select_baseline(self.current, CONTEXT, 'owner/repo', '100', self.api)
        self.assertIsNone(previous)
        self.assertIn('revision', notes[0])

    def test_unavailable_artifact_does_not_stop_searching_older_runs(self):
        self.runs = [candidate(i, f'2026-11-01T0{i}:00:00Z') for i in (1, 2)]
        self.reports = {1: archive()}
        original = self.api
        def unavailable(path, binary=False):
            if path.endswith('/artifacts/2/zip'):
                raise RuntimeError('artifact unavailable')
            return original(path, binary)
        _, url, notes = bench.select_baseline(self.current, CONTEXT, 'owner/repo', '100', unavailable)
        self.assertTrue(url.endswith('/1'))
        self.assertIn('unavailable', notes[0])


if __name__ == '__main__':
    unittest.main()
