from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from rapid_main.io.sample_index import (
    registrations_to_samples,
    read_sample_index_registrations,
)


class TestSampleIndexIO(unittest.TestCase):
    def test_read_sample_index_with_coordinate_header(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "field_a.sam"
            p.write_text(
                "\n".join(
                    [
                        "Site_A",
                        "47.2 -122.1",
                        "BK-01",
                        "BK-02",
                        "BK-03",
                    ]
                ),
                encoding="latin-1",
            )

            registrations = read_sample_index_registrations(p)
            self.assertEqual([r.specimen_name for r in registrations.entries], ["BK-01", "BK-02", "BK-03"])
            self.assertEqual(registrations.entries[0].sample_set, "Site_A")
            self.assertEqual(registrations.entries[0].location, "47.2 -122.1")
            self.assertEqual(registrations_to_samples(registrations), ["BK-01", "BK-02", "BK-03"])
            self.assertEqual(registrations.entries[0].order, 1)
            self.assertEqual(registrations.entries[0].source_file, str(p.resolve()))

    def test_csv_index_preserves_source_identity_and_sparse_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'index.csv'
            path.write_text('Sample Name,Depth (cm),Formation,Location\nA,2,Unit A,Site A\nB\n', encoding='utf-8')
            records = read_sample_index_registrations(path)
            self.assertEqual(records.names, ['A', 'B'])
            self.assertEqual(records.entries[0].formation, 'Unit A')
            self.assertEqual(records.entries[1].location, '')
            self.assertTrue(all(entry.source_file == str(path.resolve()) for entry in records.entries))

    def test_read_sample_index_plain_list(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "plain.sam"
            p.write_text(
                "BK-01\nBK-02  42.3\n#comment\nBK-03\n",
                encoding="latin-1",
            )

            registrations = read_sample_index_registrations(p)
            self.assertEqual(registrations.entries[0].sample_set, "plain")
            self.assertEqual(registrations.entries[0].depth_cm, "")
            self.assertEqual(registrations.entries[1].depth_cm, "42.3")
            self.assertEqual(registrations.entries[2].specimen_name, "BK-03")
            self.assertEqual(registrations.entries[2].location, str(p.parent))


if __name__ == "__main__":
    unittest.main(verbosity=2)
