import unittest
from rapid_main.queue_station import QueueStationGeometry


class QueueStationTests(unittest.TestCase):
    def geometry(self, **overrides):
        values = dict(calibration_source='accepted.ini', use_xy_table=False,
                      slot_min=1, slot_max=100, hole_slot=10, one_step=-1000.5)
        return QueueStationGeometry(**dict(values, **overrides))

    def test_chain_uses_imported_hole_interval_not_slot_count(self):
        geometry = self.geometry()
        self.assertTrue(geometry.is_empty(10))
        self.assertTrue(geometry.is_empty(90))
        self.assertFalse(geometry.is_empty(11))
        self.assertFalse(geometry.is_empty(0))
        self.assertFalse(geometry.is_empty(110))

    def test_xy_has_one_explicit_empty_location(self):
        geometry = self.geometry(use_xy_table=True, hole_slot=46)
        self.assertTrue(geometry.is_empty(46))
        self.assertFalse(geometry.is_empty(92))
        self.assertEqual(geometry.nearest_empty(1), 46)

    def test_chain_wrap_and_upper_tie_match_vb6(self):
        geometry = self.geometry(slot_max=101)
        for current, target in ((15, 20), (14, 10), (99, 100), (101, 100), (1, 100), (10, 10)):
            with self.subTest(current=current):
                self.assertEqual(geometry.nearest_empty(current), target)

    def test_holder_zero_resolves_to_empty_slot_and_never_specimen(self):
        geometry = self.geometry()
        self.assertEqual(geometry.resolve_holder(0, current_slot=14), 10)
        self.assertEqual(geometry.resolve_holder(20, current_slot=14), 20)
        for slot in (-1, 1, 101):
            with self.assertRaises(ValueError):
                geometry.resolve_holder(slot, current_slot=14)
        self.assertEqual(geometry.specimen_slot(14), 14)
        with self.assertRaises(ValueError):
            geometry.specimen_slot(10)

    def test_live_position_alignment_and_wrong_hole_fail_closed(self):
        geometry = self.geometry(one_step=-1000)
        self.assertEqual(geometry.verify_empty_readback(10, -10010), 10)
        for target, counts in ((10, -20000), (10, -10100), (11, -11000), (10, True), (10, 2**31)):
            with self.subTest(target=target, counts=counts), self.assertRaises(ValueError):
                geometry.verify_empty_readback(target, counts)

    def test_chain_readback_wraps_both_signs_and_fractional_calibration(self):
        geometry = self.geometry()
        for counts, expected in ((-10005, 10), (-110055, 10), (90045, 10), (0, 100)):
            self.assertEqual(geometry.slot_from_counts(counts), expected)

    def test_xy_readback_does_not_wrap_out_of_range_into_valid_hole(self):
        geometry = self.geometry(use_xy_table=True, hole_slot=46, one_step=-1000,
                                 xy_positions=((46, 0, 39), (1, 9590, -11916)))
        self.assertEqual(geometry.verify_empty_xy_readback(46, 0, 39), 46)
        with self.assertRaises(ValueError):
            geometry.verify_empty_readback(46, -46000)
        for counts in (-146000, 54000, 0):
            with self.assertRaises(ValueError):
                geometry.slot_from_counts(counts)

    def test_missing_and_ambiguous_geometry_is_rejected(self):
        for overrides in (dict(calibration_source=''), dict(use_xy_table='true'),
                          dict(slot_min=2), dict(hole_slot=0), dict(hole_slot=101),
                          dict(slot_max=0), dict(one_step=0), dict(one_step=float('nan')),
                          dict(slot_min=True), dict(hole_slot=10.2), dict(one_step=True)):
            with self.subTest(overrides=overrides), self.assertRaises(ValueError):
                self.geometry(**overrides)

    def test_invalid_slot_inputs_cannot_be_truncated(self):
        for slot in (True, 1.2, float('nan'), '10', 0, 101):
            with self.subTest(slot=slot), self.assertRaises(ValueError):
                self.geometry().nearest_empty(slot)

    def test_config_requires_full_accepted_calibration_and_imported_hole(self):
        from dataclasses import asdict
        from rapid_main.config import AppConfig
        from rapidpy_common.hardware import MotorControllerConfig
        config = AppConfig()
        with self.assertRaises(ValueError):
            QueueStationGeometry.from_config(config, use_xy_table=True)
        config.motor_station.calibration_source = 'station.ini'
        config.motor_station.controller = asdict(MotorControllerConfig())
        config.motor_station.hole_slot = 46
        config.motor_station.xy_positions = {'46': [0, 39]}
        config.motor_station.xy_home = [-3, -2]
        config.motor_station.use_xy_table = True
        geometry = QueueStationGeometry.from_config(config, use_xy_table=True)
        self.assertEqual(geometry.resolve_holder(0, current_slot=1), 46)
        del config.motor_station.controller['slot_max']
        with self.assertRaises(ValueError):
            QueueStationGeometry.from_config(config, use_xy_table=True)

    def test_all_chain_slots_match_legacy_directional_search(self):
        for maximum, interval in ((101, 10), (12, 5), (10, 10), (9, 1)):
            geometry = self.geometry(slot_max=maximum, hole_slot=interval)
            for current in range(1, maximum + 1):
                upper = next((current - 1 + distance) % maximum + 1
                             for distance in range(maximum)
                             if geometry.is_empty((current - 1 + distance) % maximum + 1))
                lower = next((current - 1 - distance) % maximum + 1
                             for distance in range(maximum)
                             if geometry.is_empty((current - 1 - distance) % maximum + 1))
                expected = upper if (upper - current) % maximum <= (current - lower) % maximum else lower
                self.assertEqual(geometry.nearest_empty(current), expected)

    def test_extreme_count_conversion_fails_with_clear_validation_error(self):
        with self.assertRaises(ValueError):
            self.geometry(one_step=5e-324).slot_from_counts(1000)

    def test_xy_map_checks_both_axes_and_never_falls_back_to_slot_min(self):
        geometry = self.geometry(use_xy_table=True, hole_slot=46,
            xy_positions=((46, 0, 39), (1, 9590, -11916)))
        self.assertEqual(geometry.xy_target(1), (9590, -11916))
        self.assertEqual(geometry.slot_from_xy_counts(9590, -11916), 1)
        for x, y in ((0, -100000), (0, 100000), (9590, 39), (True, 39), (0, 2**31)):
            with self.subTest(x=x, y=y), self.assertRaises(ValueError):
                geometry.verify_empty_xy_readback(46, x, y)
        with self.assertRaises(ValueError):
            geometry.xy_target(2)

    def test_ambiguous_and_missing_xy_positions_cannot_prove_empty_hole(self):
        for positions in ((), ((46, 0, 39), (1, 0, 39)), ((46, 0, 39), (1, 1, 40))):
            geometry = self.geometry(use_xy_table=True, hole_slot=46, xy_positions=positions)
            with self.assertRaises(ValueError):
                geometry.verify_empty_xy_readback(46, 0, 39)

    def test_xy_coordinates_are_immutable_and_strict_signed_integer_counts(self):
        original = [[46, 0, 39]]
        geometry = self.geometry(use_xy_table=True, hole_slot=46, xy_positions=original)
        original[0][1] = 500
        self.assertEqual(geometry.xy_target(46), (0, 39))
        for positions in (((46, True, 39),), ((46, 0.2, 39),), ((46, 0, 2**31),),
                          ((46, 0, 39), (46, 1, 40)), ((101, 0, 39),)):
            with self.subTest(positions=positions), self.assertRaises(ValueError):
                self.geometry(use_xy_table=True, hole_slot=46, xy_positions=positions)

    def test_station_mode_mismatch_and_incomplete_xy_calibration_fail_closed(self):
        from dataclasses import asdict
        from rapid_main.config import AppConfig
        from rapidpy_common.hardware import MotorControllerConfig
        config = AppConfig()
        config.motor_station.calibration_source = 'station.ini'
        config.motor_station.controller = asdict(MotorControllerConfig())
        config.motor_station.hole_slot = 46
        for positions, home, mode in (({}, [], True), ({'46': [0, 39]}, [], True),
                                      ({'46': [0, 39]}, [-3, -2], False),
                                      ({'046': [0, 39]}, [-3, -2], True)):
            config.motor_station.xy_positions = positions
            config.motor_station.xy_home = home
            config.motor_station.use_xy_table = mode
            with self.assertRaises(ValueError):
                QueueStationGeometry.from_config(config, use_xy_table=True)


if __name__ == '__main__':
    unittest.main()
