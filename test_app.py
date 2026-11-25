import unittest
import pandas as pd
import numpy as np
import os
from data_loader import parse_file_1, parse_file_2, link_datasets, to_clean_numeric
from app import to_excel_download_bytes

class TestSpendAnalyticsEngine(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        # Ensure sample files exist
        if not os.path.exists("1.xlsx") or not os.path.exists("2.xlsx"):
            import generate_sample_data
            generate_sample_data.create_sample_files()
            
        cls.df1, cls.map1 = parse_file_1("1.xlsx")
        cls.df2, cls.map2 = parse_file_2("2.xlsx")

    def test_01_file1_parsing_and_super_headers(self):
        """Verify File 1 parsed correctly with super-headers combined and separators dropped."""
        self.assertGreater(len(self.df1), 0, "File 1 should contain rows")
        self.assertIn("Item Code", self.df1.columns)
        self.assertIn("Item Description", self.df1.columns)
        
        # Verify separator columns ('x' or 'X') were dropped
        for col in self.df1.columns:
            self.assertNotIn(col.lower().strip(), ['x', '__drop_sep__'], f"Separator column {col} was not dropped")
            
        # Verify super-headers combined for monthly columns
        monthly_spend = self.map1.get('monthly_spend_2026', {})
        self.assertEqual(len(monthly_spend), 12, "Should have mapped all 12 months for 2026 Spend")
        
        # Verify total columns mapped
        self.assertIsNotNone(self.map1.get('spend_2026'), "Total 2026 Spend column must be mapped")
        self.assertIsNotNone(self.map1.get('spend_2025'), "Total 2025 Spend column must be mapped")
        self.assertIsNotNone(self.map1.get('vol_2026'), "Total 2026 Volume column must be mapped")
        self.assertIsNotNone(self.map1.get('ppv'), "PPV column must be mapped")
        self.assertIsNotNone(self.map1.get('fx_impact'), "Fx Impact column must be mapped")

    def test_02_file2_parsing_and_metadata_bypass(self):
        """Verify File 2 bypassed top metadata rows and detected 'Item Name' header."""
        self.assertGreater(len(self.df2), 0, "File 2 should contain rows")
        self.assertIn("Item Name", self.df2.columns)
        self.assertIn("Price", self.df2.columns)
        self.assertIn("Agreement Number", self.df2.columns)
        self.assertIn("Supplier", self.df2.columns)

    def test_03_linking_inner_and_left_join(self):
        """Verify linking logic and join toggles."""
        linked_left = link_datasets(self.df1, self.df2, self.map1, self.map2, join_type="left")
        linked_inner = link_datasets(self.df1, self.df2, self.map1, self.map2, join_type="inner")
        
        stats_left = linked_left['stats']
        stats_inner = linked_inner['stats']
        
        self.assertEqual(stats_left['total_items_f1'], 25)
        self.assertEqual(len(linked_left['item_df']), 25, "Left join should keep all 25 items from File 1")
        self.assertEqual(len(linked_inner['item_df']), stats_left['matched_items_count'], "Inner join should only keep matched items")
        self.assertLess(len(linked_inner['item_df']), len(linked_left['item_df']))

    def test_04_double_counting_prevention(self):
        """Verify File 1 spend & volume are NOT double-counted despite 1-to-many agreements in File 2."""
        linked = link_datasets(self.df1, self.df2, self.map1, self.map2, join_type="left")
        
        item_df = linked['item_df']
        agreement_df = linked['agreement_df']
        
        spend_col = self.map1['spend_2026']
        vol_col = self.map1['vol_2026']
        
        # Original sum in File 1
        expected_spend = self.df1[spend_col].sum()
        expected_vol = self.df1[vol_col].sum()
        
        # Item-level joined sum
        actual_spend = item_df[spend_col].sum()
        actual_vol = item_df[vol_col].sum()
        
        self.assertAlmostEqual(actual_spend, expected_spend, places=2, 
                               msg="Spend in item_df must match original File 1 spend exactly (no inflation!)")
        self.assertAlmostEqual(actual_vol, expected_vol, places=2, 
                               msg="Volume in item_df must match original File 1 volume exactly (no inflation!)")
        
        # Verify agreement_df has more rows than item_df due to multiple contracts
        self.assertGreater(len(agreement_df), len(item_df), 
                           "Agreement table should contain detailed 1-to-many lines")

    def test_05_active_agreements_filter_toggle(self):
        """Verify active-only agreements toggle."""
        linked_all = link_datasets(self.df1, self.df2, self.map1, self.map2, active_only=False)
        linked_active = link_datasets(self.df1, self.df2, self.map1, self.map2, active_only=True)
        
        # Agreement rows should be fewer or equal when active_only is True
        self.assertLessEqual(len(linked_active['agreement_df']), len(linked_all['agreement_df']))

    def test_06_excel_export_generator(self):
        """Verify clean Excel export generator."""
        sample_df = self.df1[['Item Code', 'Brand', 'Category']].head(5)
        excel_bytes = to_excel_download_bytes(sample_df, "TestSheet")
        self.assertIsInstance(excel_bytes, bytes)
        self.assertGreater(len(excel_bytes), 100)

if __name__ == '__main__':
    unittest.main()

