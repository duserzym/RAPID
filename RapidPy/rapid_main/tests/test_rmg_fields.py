"""RMG row compatibility and the evidence-backed measurement export."""

import csv
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest

from rapid_main.af_treatment import AfPass, AfTreatmentPlan, AfTreatmentRecord, CalibratedAfRamp
from rapid_main.data_model import SpecimenMeta
from rapid_main.io.rmg_reader import read_rmg
from rapid_main.io.rmg_writer import describe_rmg_treatment, write_rmg_records
from rapid_main.measurement_worker import MeasurementWorker
from tests.test_specimen_orientation import BracketedStation


class RmgFieldsTests(unittest.TestCase):
    def test_reads_a_legacy_header_data_row_and_sparse_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'VB6.rmg'
            source.write_text(
                'SPEC,original,Vol: 8.2, \n'
                ' ,Level,Bias Field (G),Spin Speed (rps),Hold Time (s),Mz (emu),Std. Dev. Z,Mz/Vol, '
                'Moment Susceptibility (emu/Oe),Mx (emu),Std. Dev. X,My (emu),Std. Dev. Y,Remarks,'
                'Core Dec,Core Inc,M (emu),CSD,Sample Height (cm),Date/Time\n'
                'RRMz,600,30,-5,1.2,0.01,0.001,0.00122,0.012,0.02,0.002,-0.03,0.003,'
                'original,310,-25,0.0045,2.1,2.8,11/22/2026 10:00:00 AM \n'
                'AF,100,,,,,,,0.02\n', encoding='latin-1')
            first, old = read_rmg(source)
            self.assertEqual((first.step_type, first.step_value, first.spin_speed_rps), ('RRMz', 600., -5.))
            self.assertEqual((first.moment_y_emu, first.sd_y_emu, first.core_inc), (-.03, .003, -25.))
            self.assertEqual((first.susceptibility, first.intensity_emu_per_cm3, first.sample_height_cm), (.012, .0045, 2.8))
            self.assertEqual((old.step_type, old.susceptibility, old.sd_x_emu), ('AF', .02, None))
            target = root / 'rewritten.rmg'
            write_rmg_records(target, [first, old])
            self.assertEqual(read_rmg(target), [first, old])

    def test_treatment_units_and_matching_verified_pulse_axis(self):
        self.assertEqual(describe_rmg_treatment('AF20').level, 200.)
        self.assertEqual(describe_rmg_treatment('IRMZ-1200G').level, -1200.)
        self.assertEqual(describe_rmg_treatment('TT350').level, 350.)
        self.assertIsNone(describe_rmg_treatment('AFmax').level)
        rrm = describe_rmg_treatment('RRMZ60/-5@3')
        self.assertEqual((rrm.step_type, rrm.level, rrm.spin_speed_rps, rrm.bias_field_g), ('RRMz', 600., -5., 30.))
        pulse = SimpleNamespace(treatment_id='pulse', circuit=SimpleNamespace(plan=SimpleNamespace(
            coil='transverse', field_mT=-80.)))
        irm = describe_rmg_treatment('IRM-80', pulse=pulse)
        self.assertEqual((irm.step_type, irm.level, irm.evidence_id), ('IRMx', -800., 'pulse'))

    def test_worker_uses_this_steps_af_record_and_keeps_original_units(self):
        class Station(BracketedStation):
            def __init__(self):
                super().__init__()
                self.af_treatment_records = []

            def set_demag_step(self, label):
                # A measured ARM run can supply bias and AF values that are
                # absent from its nominal label, through the accepted record.
                ramp = CalibratedAfRamp(label, 'axial', 23., 1., 1., 50., 1., 1., 1250, 100., 'fixture')
                plan = AfTreatmentPlan(label, -100, (AfPass(0., ramp),))
                self.af_treatment_records.append(AfTreatmentRecord('af-' + 'a' * 32, 'A', 'rmg-fixture',
                    plan, 1, (), '', (), True, True, bias_mT=1.25))

            def measured_specimen_height_cm(self):
                return 2.8

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            worker = MeasurementWorker(SpecimenMeta('A', comment='quoted, comment', volume=8.2),
                                       ['ARM'], output, Station(), run_id='rmg-fixture')
            finished = []
            worker.run_finished.connect(finished.append)
            worker.run()
            self.assertEqual(finished, [False], worker._error_messages)

            published = output / 'SIMULATED' / 'A.rmg'
            with published.open(encoding='latin-1', newline='') as handle:
                header, columns, row = list(csv.reader(handle))
            self.assertEqual(header[1], 'quoted, comment')
            self.assertEqual(len(row), 20)
            self.assertEqual(columns[8], 'Moment Susceptibility (emu/Oe)')
            record, = read_rmg(published)
            self.assertEqual((record.step_value, record.bias_field_g, record.hold_time_s), (230., 12.5, 1.25))
            self.assertAlmostEqual(record.moment_z_emu, 1e-5)
            self.assertAlmostEqual(record.intensity_z_emu_per_cm3, 1e-5 / 8.2)
            self.assertAlmostEqual(record.sd_x_emu, (.16 / 3.) ** .5 * 1e-5)
            self.assertEqual(record.sample_height_cm, 2.8)
            captured, = worker._progress.accepted_steps()
            self.assertEqual(captured['rmg_treatment']['evidence_id'], 'af-' + 'a' * 32)
            self.assertEqual(captured['sample_height_cm'], 2.8)
