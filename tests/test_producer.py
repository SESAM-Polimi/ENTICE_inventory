"""Producer integration checks: ENTICE_PRODUCER_TESTS=1 in the pinned environment."""
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest


@unittest.skipUnless(os.environ.get("ENTICE_PRODUCER_TESTS") == "1", "Requires the pinned producer environment; set ENTICE_PRODUCER_TESTS=1")
class ProducerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import mario
        import pandas as pd
        cls.mario, cls.pd = mario, pd
        path = Path(__file__).resolve().parents[1] / "scripts/verify_producer.py"
        spec = importlib.util.spec_from_file_location("producer_verification", path)
        cls.verify = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cls.verify)

    def fixture(self):
        pd = self.pd
        index = pd.MultiIndex.from_product([["R1", "R2"], ["Sector"], ["Parent", "Stable"]], names=["Region", "Level", "Item"])
        demand = pd.MultiIndex.from_product([["R1", "R2"], ["Consumption category"], ["Households"]], names=index.names)
        z = pd.DataFrame([[20, 15, 4, 5], [10, 54, 5, 5], [5, 5, 18, 11], [3, 5, 9, 41]], index=index, columns=index, dtype=float)
        y = pd.DataFrame([[15, 25], [20, 45], [15, 30], [35, 30]], index=index, columns=demand, dtype=float)
        v = pd.DataFrame([[46, 60, 48, 61]], index=pd.Index(["Value added"], name="Item"), columns=index, dtype=float)
        e = pd.DataFrame([[10, 5, 20, 10]], index=pd.Index(["CO2"], name="Item"), columns=index, dtype=float)
        ey = pd.DataFrame([[4, 2]], index=e.index, columns=demand, dtype=float)
        units = {kind: pd.DataFrame({"unit": values}, index=labels) for kind, labels, values in [
            ("Sector", ["Parent", "Stable"], ["M USD", "M USD"]),
            ("Factor of production", ["Value added"], ["M USD"]),
            ("Satellite account", ["CO2"], ["tonne"]),
        ]}
        return self.mario.Database(name="Exactly balanced synthetic fixture", table="IOT", Z=z, Y=y, V=v, E=e, EY=ey, units=units)

    def test_metric_rejects_missing_and_non_finite_entries(self):
        actual = self.pd.Series([1.0, float("nan")], index=["a", "b"])
        expected = self.pd.Series([1.0, 2.0], index=["a", "c"])
        check = self.verify.metric(actual, expected, 1e-6, 1e-8)
        self.assertFalse(check["passed"])
        self.assertEqual(check["non_finite_entries"], 2)
        self.assertEqual({row["row"] for row in check["failure_examples"]}, {"b", "c"})
        json.dumps(check, allow_nan=False)

    def test_chunked_aggregation_preserves_flows(self):
        db = self.fixture()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "Z.parquet"
            db.Z.to_parquet(path)
            actual = self.verify.aggregate_parquet(path, "R1", chunk_size=1)
            expected = db.Z.rename(index={"R2": "ROW"}, columns={"R2": "ROW"}).sort_index().sort_index(axis=1)
            self.pd.testing.assert_frame_equal(actual, expected)

    def test_baseline_diagnostics_do_not_correct_source_residuals(self):
        db = self.fixture()
        db.update_scenarios("baseline", V=db.V * 1.00001)
        original = db.V.copy(deep=True)
        self.assertFalse(self.verify.accounting(db, 1e-8, 1e-10)["column_accounting"]["passed"])
        self.pd.testing.assert_frame_equal(db.V, original)
        self.assertFalse(hasattr(self.verify, "reconcile_baseline"))

    def test_split_preserves_accounting_parent_totals_and_observed_trade(self):
        from mario.ops.add_sector_specs import (
            ADD_SECTOR_SPLIT_EXCLUSION_COLUMNS, ADD_SECTOR_SPLIT_OUTPUT_COLUMNS,
            ADD_SECTOR_SPLIT_TRADE_COLUMNS,
        )
        pd = self.pd
        db = self.fixture()
        region, parent, child = "R1", "Parent", "Child"
        column = (region, "Sector", parent)
        quantity = float(db.X.loc[column].iloc[0]) * 0.1
        trade = float(db.Z.loc[column, "R2"].sum() + db.Y.loc[column, "R2"].sum()) * 0.1
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            workbook = folder / "input.xlsx"
            db.get_add_sectors_excel(items=[child], regions=[region], path=workbook)
            master = pd.read_excel(workbook, sheet_name="Master").astype(object)
            master.loc[0, "Unit"] = "M USD"
            master.loc[0, "Parent Sector"] = parent
            master.loc[0, "Add or Split"] = "Split"
            rows = []
            for (r, kind, item), value in db.z[column].items():
                rows.append([value, "M USD", item, kind, item, r, "Update"])
            rows.append([float(db.v[column].iloc[0]), "M USD", "Value added", "Factor of production", "Value added", "", "Update"])
            frames = {
                "Master": master,
                "INV_001": pd.DataFrame(rows, columns=["Quantity", "Unit", "Input", "Item type", "DB Item", "DB Region", "Change type"]),
                "Total outputs": pd.DataFrame([[child, region, quantity, "M USD", "synthetic", ""]], columns=ADD_SECTOR_SPLIT_OUTPUT_COLUMNS.values()),
                "Trades": pd.DataFrame([[child, region, "R2", trade, "M USD", "synthetic", ""]], columns=ADD_SECTOR_SPLIT_TRADE_COLUMNS.values()),
                "Exclusions": pd.DataFrame(columns=ADD_SECTOR_SPLIT_EXCLUSION_COLUMNS.values()),
                "Tolerances": pd.DataFrame({"tol_Name": ["delta", "eps"], "values": [0.0, 0.0]}),
            }
            with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                for name, frame in frames.items():
                    frame.to_excel(writer, sheet_name=name, index=False)
            db.read_add_sectors_excel(workbook, read_inventories=True, split=True)
            result = db.add_sectors(inplace=False, split=True, solver="CLARABEL", cvxlab_path=folder,
                                    VA_fix=True, accept_non_unitary_sum=True, ignore_warnings=False,
                                    solver_parameters={"tol_gap_abs": 1e-9, "tol_gap_rel": 1e-9, "tol_feas": 1e-9})
            checks = self.verify.accounting(result, 1e-6, 1e-8)
            for name in ("row_accounting", "column_accounting"):
                self.assertTrue(checks[name]["passed"], checks[name])
            for matrix in ("Z", "Y", "V"):
                check = self.verify.metric(self.verify.collapsed(getattr(result, matrix), child, parent), getattr(db, matrix), 1e-6, 1e-8)
                self.assertTrue(check["passed"], (matrix, check))
            supplied = float(result.Z.loc[(region, "Sector", child), "R2"].sum() + result.Y.loc[(region, "Sector", child), "R2"].sum())
            self.assertAlmostEqual(supplied, trade, places=6)
            self.assertAlmostEqual(float(result.X.loc[(region, "Sector", child)].iloc[0]), quantity, places=6)


if __name__ == "__main__":
    unittest.main()
