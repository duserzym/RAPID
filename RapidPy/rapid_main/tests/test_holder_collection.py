from dataclasses import replace
import hashlib
import json
import math
from pathlib import Path
import tempfile
import unittest

from rapid_main.holder_collection import canonical
from rapid_main.holder_measurement import HolderMeasurementService
from rapid_main.holder_state import HolderCorrection, HolderMetrics, HolderStateError, HolderStateStore
from rapid_main.magnetometer import BlockAudit, BracketedMeasurementBlock, reduce_bracketed_measurement

NOW = '2026-10-04T12:00:00+00:00'


def source(identity, scale=1., drift=0.):
    return BracketedMeasurementBlock((0., 0., 0.),
        tuple((2*scale + drift*(index+1)/5, 0., scale) for index in range(4)), (drift, 0., 0.),
        audit=BlockAudit(block_id=identity, sample_name='Holder', run_id='run',
                         config_hash='config', is_holder_block=True))


def correction(blocks=None):
    return HolderCorrection.from_collection(blocks or (source('one'), source('two', 2)),
        holder_id='HOLDER', measured_at_iso=NOW)


class HolderCollectionTests(unittest.TestCase):
    def test_quality_uses_all_eight_positions_and_mean_induced_and_drift(self):
        value = correction((source('one', 1, .005), source('two', 2, .015)))
        self.assertAlmostEqual(value.metrics.magnitude_raw, 1.5)
        self.assertAlmostEqual(value.metrics.induced_magnitude_raw, 3)
        self.assertAlmostEqual(value.metrics.drift_magnitude_raw, .01)
        self.assertAlmostEqual(value.metrics.asymmetry_ratio, 2)
        self.assertAlmostEqual(value.metrics.fischer_sd_deg, 81*math.sqrt((8-8/math.sqrt(5))/7))
        self.assertEqual(value.averaging_cycles, 2)

    def test_calibrated_collection_moment_respects_anisotropic_gains(self):
        block = replace(source('one'), axis_calibration=(2., 3., 4.))
        value = correction((block,))
        self.assertAlmostEqual(value.metrics.magnitude_emu, 4e-5)

    def test_all_raw_blocks_audits_and_aggregate_statistics_survive_round_trip(self):
        original = correction()
        packet = json.loads(original.collection_evidence_json)
        self.assertEqual([raw['audit']['block_id'] for raw in packet['blocks']], ['one', 'two'])
        self.assertEqual(packet['blocks'][1]['positions'][0], [4, 0, 2])
        self.assertEqual(packet['statistics']['position_count'], 8)
        self.assertEqual(hashlib.sha256(original.collection_evidence_json.encode()).hexdigest(), original.collection_sha256)
        self.assertIn(original.collection_sha256, original.record_version)
        loaded = HolderCorrection.from_dict(json.loads(json.dumps(original.to_dict())))
        loaded.verify_collection()
        self.assertEqual(loaded, original)

    def test_same_holder_and_timestamp_with_different_sources_get_distinct_versions(self):
        first = correction()
        second = correction((source('changed'), source('two', 2)))
        self.assertNotEqual(first.record_version, second.record_version)

    def test_caller_mutable_raw_containers_cannot_edit_accepted_correction(self):
        raw = source('one')
        positions = [list(vector) for vector in raw.positions]
        raw = replace(raw, positions=positions)
        value = correction((raw,))
        before = value.collection_evidence_json, value.positions
        positions[0][0] = 999
        value.verify_collection()
        self.assertEqual((value.collection_evidence_json, value.positions), before)

    def test_tampered_blob_and_outer_correction_never_replace_previous(self):
        previous = correction()
        store = HolderStateStore()
        store.install(previous)
        packet = json.loads(previous.collection_evidence_json)
        packet['blocks'][0]['positions'][0][0] += 1
        changed = canonical(packet)
        for bad in (replace(previous, collection_evidence_json=changed),
                    replace(previous, positions=((9., 9., 9.),) * 4),
                    replace(previous, averaging_cycles=3),
                    replace(previous, metrics=replace(previous.metrics, fischer_sd_deg=1.))):
            with self.subTest(bad=bad.collection_sha256), self.assertRaises(HolderStateError):
                store.install(bad)
            self.assertIs(store.current, previous)

    def test_rehashed_but_inconsistent_aggregate_packet_is_rejected(self):
        original = correction()
        packet = json.loads(original.collection_evidence_json)
        packet['statistics']['mean_raw'][2] = 999
        text = canonical(packet)
        forged = replace(original, collection_evidence_json=text,
                         collection_sha256=hashlib.sha256(text.encode()).hexdigest())
        with self.assertRaisesRegex(HolderStateError, 'aggregates disagree'):
            forged.verify_collection()

    def test_legacy_single_block_remains_readable_but_unauditable_multi_block_requires_remeasurement(self):
        single = HolderCorrection.from_result(reduce_bracketed_measurement(source('one')),
            holder_id='OLD', measured_at_iso=NOW)
        self.assertTrue(single.is_finite())
        old_multi = replace(single, averaging_cycles=2)
        with self.assertRaisesRegex(HolderStateError, 'remeasure'):
            HolderStateStore().install(old_multi)

    def test_bad_source_context_identity_direction_and_nonblank_correction_retain_previous(self):
        previous = correction()
        for bad in (source('one'), replace(source('two'), range_factor=2e-5),
                    replace(source('two'), is_up=False),
                    replace(source('two'), holder_positions=((1., 0., 0.),) * 4),
                    replace(source('two'), audit=replace(source('two').audit, config_hash='different')),
                    replace(source('two'), audit=replace(source('two').audit, is_holder_block=False))):
            store = HolderStateStore()
            store.install(previous)
            blocks = iter((source('one'), bad))
            result = HolderMeasurementService(lambda: next(blocks), store, averaging_cycles=2).measure(holder_id='NEW')
            self.assertFalse(result.installed)
            self.assertIs(store.current, previous)

    def test_failed_persisted_reload_preserves_current_and_good_packet_reloads(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'holder.json'
            store = HolderStateStore(path)
            original = correction()
            store.install(original)
            self.assertEqual(HolderStateStore(path).load(), original)
            packet = json.loads(path.read_text())
            packet['collection_sha256'] = '0' * 64
            path.write_text(json.dumps(packet))
            self.assertIs(store.load(), original)

    def test_nonfinite_legacy_quality_is_not_installable(self):
        old = HolderCorrection('OLD', measured_at_iso=NOW, metrics=HolderMetrics(fischer_sd_deg=float('nan')))
        with self.assertRaises(HolderStateError):
            HolderStateStore().install(old)


if __name__ == '__main__':
    unittest.main()
