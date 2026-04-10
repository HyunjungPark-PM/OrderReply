"""
Excel file comparison and processing module for p-net order reply system.
"""
import pandas as pd
from itertools import combinations
from typing import Any, List, Dict, Optional


class ExcelProcessor:
    """Handle Excel file reading, comparison, and output generation."""

    def __init__(self):
        self.pnet_download = None
        self.factory_reply = None
        self.result = None

    def read_pnet_download(self, file_path: str) -> bool:
        """
        Read p-net download file.

        Expected columns:
        - E: CPO#
        - F: CPO-LINE#
        - G: LINE SEQ
        - H: CPO QTY
        - I: Material
        - O: EX-F
        - P: ETD
        - X: PO#
        - Y: PO-LINE#
        """
        try:
            df = pd.read_excel(file_path, header=0, dtype=object)
            selected = df.iloc[:, [4, 5, 6, 7, 8, 14, 15, 23, 24]].copy()
            selected.columns = [
                'CPO#', 'CPO-LINE#', 'LINE SEQ', 'CPO QTY', 'Material',
                'EX-F', 'ETD', 'PO#', 'PO-LINE#'
            ]
            selected['CPO QTY'] = pd.to_numeric(selected['CPO QTY'], errors='coerce')
            selected['PO#'] = selected['PO#'].astype(str).str.strip()
            selected['PO-LINE#'] = selected['PO-LINE#'].astype(str).str.strip()
            selected['CPO#'] = selected['CPO#'].astype(str).str.strip()
            selected['CPO-LINE#'] = selected['CPO-LINE#'].astype(str).str.strip()
            selected['LINE SEQ'] = selected['LINE SEQ'].astype(str).str.strip()
            self.pnet_download = selected.dropna(subset=['PO#', 'PO-LINE#']).reset_index(drop=True)
            return True
        except Exception as e:
            print(f"Error reading p-net download file: {e}")
            return False

    def read_factory_reply(self, file_path: str) -> bool:
        """
        Read factory reply file.

        Expected columns:
        - A: PO#
        - B: LINE#
        - C: Material
        - D: CPO QTY
        - E: ETD
        - F: EX-F
        - G: 내부노트
        """
        try:
            df = pd.read_excel(file_path, header=0, dtype=object)
            selected = df.iloc[:, :7].copy()
            selected.columns = [
                'PO#', 'LINE#', 'Material', 'CPO QTY', 'ETD', 'EX-F', '내부노트'
            ]
            selected['CPO QTY'] = pd.to_numeric(selected['CPO QTY'], errors='coerce')
            selected['PO#'] = selected['PO#'].astype(str).str.strip()
            selected['LINE#'] = selected['LINE#'].astype(str).str.strip()
            self.factory_reply = selected.dropna(subset=['PO#', 'LINE#']).reset_index(drop=True)
            return True
        except Exception as e:
            print(f"Error reading factory reply file: {e}")
            return False

    def compare_and_generate(self) -> bool:
        """
        Compare two files and generate output file.

        Comparison key: PO# + LINE#
        Output columns: PO#, PO-LINE#, Material, CPO QTY, ETD, EX-F, 내부노트, CPO#, CPO-LINE#, LINE SEQ
        """
        try:
            if self.pnet_download is None or self.factory_reply is None:
                raise ValueError("Both input files must be loaded first")

            pnet_df = self.pnet_download.copy()
            factory_df = self.factory_reply.copy()

            pnet_df['PO#'] = pnet_df['PO#'].astype(str)
            pnet_df['PO-LINE#'] = pnet_df['PO-LINE#'].astype(str)
            factory_df['PO#'] = factory_df['PO#'].astype(str)
            factory_df['LINE#'] = factory_df['LINE#'].astype(str)

            result_list: List[Dict[str, Any]] = []

            for (po_num, line_num), pnet_group in pnet_df.groupby(['PO#', 'PO-LINE#'], sort=False):
                factory_group = factory_df[
                    (factory_df['PO#'] == po_num) &
                    (factory_df['LINE#'] == line_num)
                ].copy()

                if factory_group.empty:
                    for _, pnet_row in pnet_group.iterrows():
                        result_list.append({
                            'PO#': po_num,
                            'PO-LINE#': line_num,
                            'Material': pnet_row['Material'],
                            'CPO QTY': pnet_row['CPO QTY'],
                            'ETD': pnet_row['ETD'],
                            'EX-F': pnet_row['EX-F'],
                            '내부노트': None,
                            'CPO#': pnet_row['CPO#'],
                            'CPO-LINE#': pnet_row['CPO-LINE#'],
                            'LINE SEQ': pnet_row['LINE SEQ']
                        })
                else:
                    result_list.extend(self._map_factory_to_download(pnet_group, factory_group))

            self.result = pd.DataFrame(result_list)
            return True
        except Exception as e:
            print(f"Error comparing files: {e}")
            return False

    def _map_factory_to_download(
        self,
        download_group: pd.DataFrame,
        factory_group: pd.DataFrame
    ) -> List[Dict[str, Any]]:
        """Map factory ETD/EX-F values to download rows with minimal splitting."""
        download_rows = download_group.copy().reset_index(drop=True).to_dict('records')
        for index, row in enumerate(download_rows):
            row['_orig_index'] = index
            row['_quantity'] = float(row['CPO QTY'] or 0)

        factory_rows = []
        for index, row in enumerate(factory_group.copy().reset_index(drop=True).to_dict('records')):
            remaining_qty = float(row['CPO QTY'] or 0)
            if remaining_qty > 0:
                factory_rows.append({**row, 'remaining_qty': remaining_qty, '_index': index})

        result_rows: List[Dict[str, Any]] = []

        download_rows.sort(key=lambda row: (-row['_quantity'], row['_orig_index']))

        whole_assigned = []
        split_candidates = []

        for download_row in download_rows:
            if download_row['_quantity'] <= 0:
                result_rows.append({
                    'PO#': download_row['PO#'],
                    'PO-LINE#': download_row['PO-LINE#'],
                    'Material': download_row['Material'],
                    'CPO QTY': download_row['CPO QTY'],
                    'ETD': download_row['ETD'],
                    'EX-F': download_row['EX-F'],
                    '내부노트': None,
                    'CPO#': download_row['CPO#'],
                    'CPO-LINE#': download_row['CPO-LINE#'],
                    'LINE SEQ': download_row['LINE SEQ']
                })
                continue

            candidate = self._find_whole_assignment(download_row, factory_rows)
            if candidate is not None:
                factory_row = candidate
                self._assign_factory_segment(download_row, factory_row, download_row['_quantity'], result_rows)
                whole_assigned.append(download_row)
            else:
                split_candidates.append(download_row)

        for download_row in split_candidates:
            self._assign_split_download(download_row, factory_rows, result_rows)

        result_rows.sort(key=lambda row: (row['PO#'], row['PO-LINE#'], row['LINE SEQ']))
        return result_rows

    def _find_whole_assignment(self, download_row: Dict[str, Any], factory_rows: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        download_qty = download_row['_quantity']
        exact_same_etd: List[Dict[str, Any]] = []
        same_etd: List[Dict[str, Any]] = []
        candidates: List[Dict[str, Any]] = []

        for factory_row in factory_rows:
            remaining_qty = factory_row['remaining_qty']
            if remaining_qty >= download_qty:
                candidates.append(factory_row)
                if self._matches_etd(download_row, factory_row):
                    same_etd.append(factory_row)
                    if abs(remaining_qty - download_qty) < 1e-9:
                        exact_same_etd.append(factory_row)

        if exact_same_etd:
            return min(exact_same_etd, key=lambda row: row['remaining_qty'])
        if same_etd:
            return min(same_etd, key=lambda row: row['remaining_qty'])
        if candidates:
            return min(candidates, key=lambda row: row['remaining_qty'])
        return None

    def _matches_etd(self, download_row: Dict[str, Any], factory_row: Dict[str, Any]) -> bool:
        return str(download_row['ETD']).strip() == str(factory_row['ETD']).strip()

    def _assign_factory_segment(
        self,
        download_row: Dict[str, Any],
        factory_row: Dict[str, Any],
        quantity: float,
        result_rows: List[Dict[str, Any]]
    ) -> None:
        factory_row['remaining_qty'] -= quantity
        result_rows.append({
            'PO#': download_row['PO#'],
            'PO-LINE#': download_row['PO-LINE#'],
            'Material': factory_row['Material'],
            'CPO QTY': quantity,
            'ETD': factory_row['ETD'],
            'EX-F': factory_row['EX-F'],
            '내부노트': factory_row['내부노트'],
            'CPO#': download_row['CPO#'],
            'CPO-LINE#': download_row['CPO-LINE#'],
            'LINE SEQ': download_row['LINE SEQ']
        })

    def _assign_split_download(
        self,
        download_row: Dict[str, Any],
        factory_rows: List[Dict[str, Any]],
        result_rows: List[Dict[str, Any]]
    ) -> None:
        quantity_to_assign = download_row['_quantity']
        combo = self._find_split_combo(download_row, factory_rows)
        if combo is None:
            raise ValueError(
                f"Unable to fully map download row for PO#: {download_row['PO#']}, "
                f"PO-LINE#: {download_row['PO-LINE#']}, LINE SEQ: {download_row['LINE SEQ']}"
            )

        for factory_row in combo:
            if quantity_to_assign <= 0:
                break
            assign_qty = min(factory_row['remaining_qty'], quantity_to_assign)
            self._assign_factory_segment(download_row, factory_row, assign_qty, result_rows)
            quantity_to_assign -= assign_qty

        if quantity_to_assign > 0:
            raise ValueError(
                f"Unable to fully map download row for PO#: {download_row['PO#']}, "
                f"PO-LINE#: {download_row['PO-LINE#']}, LINE SEQ: {download_row['LINE SEQ']}"
            )

    def _find_split_combo(
        self,
        download_row: Dict[str, Any],
        factory_rows: List[Dict[str, Any]]
    ) -> Optional[List[Dict[str, Any]]]:
        required_qty = download_row['_quantity']
        available = [row for row in factory_rows if row['remaining_qty'] > 0]
        available.sort(key=lambda row: -row['remaining_qty'])

        for combo_size in (2, 3):
            best_combo = None
            best_score = None
            for combo in combinations(available, combo_size):
                total_qty = sum(row['remaining_qty'] for row in combo)
                if total_qty < required_qty:
                    continue
                assigned_qty = []
                remaining = required_qty
                same_etd_qty = 0.0
                for row in combo:
                    qty = min(row['remaining_qty'], remaining)
                    assigned_qty.append(qty)
                    if self._matches_etd(download_row, row):
                        same_etd_qty += qty
                    remaining -= qty
                if remaining > 1e-9:
                    continue

                score = (
                    sum(assigned_qty),
                    same_etd_qty,
                    tuple(assigned_qty),
                    tuple(row['remaining_qty'] for row in combo)
                )
                if best_score is None or score > best_score:
                    best_score = score
                    best_combo = combo
            if best_combo is not None:
                return list(best_combo)

        return None

    def save_result(self, output_path: str) -> bool:
        """Save result to Excel file."""
        try:
            if self.result is None:
                raise ValueError("No result to save. Run compare_and_generate first.")

            column_order = [
                'PO#', 'PO-LINE#', 'Material', 'CPO QTY', 'ETD', 'EX-F',
                '내부노트', 'CPO#', 'CPO-LINE#', 'LINE SEQ'
            ]
            self.result[column_order].to_excel(output_path, index=False)
            return True
        except Exception as e:
            print(f"Error saving result: {e}")
            return False
