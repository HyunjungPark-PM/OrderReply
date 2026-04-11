import pandas as pd
from itertools import combinations
from typing import Any, List, Dict, Optional
from datetime import datetime
import openpyxl


class ExcelProcessor:
    """Handle Excel file reading, comparison, and output generation."""

    def __init__(self):
        self.pnet_download = None
        self.factory_reply = None
        self.result = None
        self.confirmation_needed = None
        self.change_summary = None

    def _date_to_yyyymmdd(self, date_val):
        """Convert date to YYYYMMDD text format."""
        if pd.isna(date_val):
            return None
        
        if isinstance(date_val, str):
            date_val = date_val.strip()
            if len(date_val) == 8 and date_val.isdigit():
                return date_val
            try:
                parsed = pd.to_datetime(date_val)
                return parsed.strftime('%Y%m%d')
            except:
                return date_val

        try:
            num_val = int(float(date_val))
            if 19000101 <= num_val <= 29991231 and len(str(num_val)) == 8:
                return str(num_val)
            parsed = pd.to_datetime(num_val, unit='D', origin=pd.Timestamp("1899-12-30"))
            return parsed.strftime('%Y%m%d')
        except:
            return str(date_val)

    def _normalize_date_for_change_summary(self, date_val):
        """Normalize any date-like value to YYYYMMDD for comparison."""
        if pd.isna(date_val):
            return None
        try:
            normalized = pd.to_datetime(date_val)
            return normalized.strftime('%Y%m%d')
        except:
            text = str(date_val).strip()
            if len(text) == 8 and text.isdigit():
                return text
            return text

    def _format_quantity_for_change_summary(self, qty):
        """Format quantity for change summary using integer if whole number."""
        try:
            num = float(qty)
            if abs(num - int(num)) < 1e-9:
                return str(int(num))
            return str(num)
        except Exception:
            return str(qty)

    def read_pnet_download(self, file_path: str) -> bool:
        """
        Read p-net download file.

        Expected columns:
        - E: CPO#
        - F: CPO-LINE#
        - G: LINE SEQ
        - H: CPO QTY
        - I: Material
        - O: EX-F (date YYYY-MM-DD)
        - P: ETD (date YYYY-MM-DD)
        - X: PO# (9 or 10 digits, pad with 0 if 9)
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
            # Pad PO# to 10 digits
            selected['PO#'] = selected['PO#'].apply(lambda x: x.zfill(10) if len(x) == 9 else x)
            # Convert dates to YYYYMMDD text
            selected['EX-F'] = selected['EX-F'].apply(self._date_to_yyyymmdd)
            selected['ETD'] = selected['ETD'].apply(self._date_to_yyyymmdd)
            self.pnet_download = selected.dropna(subset=['PO#', 'PO-LINE#']).reset_index(drop=True)
            return True
        except Exception as e:
            print(f"Error reading p-net download file: {e}")
            return False

    def read_factory_reply(self, file_path: str) -> bool:
        """
        Read factory reply file.

        Expected columns:
        - A: PO# (9 or 10 digits, pad with 0 if 9)
        - B: LINE#
        - C: Material
        - D: CPO QTY
        - E: ETD (date YYYY-MM-DD or text YYYYMMDD)
        - F: EX-F (date YYYY-MM-DD or text YYYYMMDD)
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
            # Ensure ETD and EX-F are strings before date processing
            selected['ETD'] = selected['ETD'].astype(str).str.strip()
            selected['EX-F'] = selected['EX-F'].astype(str).str.strip()
            # Pad PO# to 10 digits
            selected['PO#'] = selected['PO#'].apply(lambda x: x.zfill(10) if len(x) == 9 else x)
            # Convert dates to YYYYMMDD text for comparison
            selected['ETD'] = selected['ETD'].apply(self._date_to_yyyymmdd)
            selected['EX-F'] = selected['EX-F'].apply(self._date_to_yyyymmdd)
            self.factory_reply = selected.dropna(subset=['PO#', 'LINE#']).reset_index(drop=True)
            return True
        except Exception as e:
            print(f"Error reading factory reply file: {e}")
            return False

    def compare_and_generate(self) -> bool:
        """
        Compare two files and generate output file with multiple sheets.

        Comparison key: PO# + LINE# (factory LINE#)
        Output sheets: 수동 업로드, 확인필요, 변경요약
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

            upload_result_list: List[Dict[str, Any]] = []
            confirmation_list: List[Dict[str, Any]] = []

            for (po_num, line_num), pnet_group in pnet_df.groupby(['PO#', 'PO-LINE#'], sort=False):
                factory_group = factory_df[
                    (factory_df['PO#'] == po_num) &
                    (factory_df['LINE#'] == line_num)
                ].copy()

                if factory_group.empty:
                    # No matching factory data, exclude
                    confirmation_list.append({
                        'PO#': po_num,
                        'PO-LINE#': line_num,
                        '사유': '공장납기회신 파일에 해당 PO#/LINE# 없음'
                    })
                    continue

                # Total quantity check
                pnet_total = pnet_group['CPO QTY'].sum()
                factory_total = factory_group['CPO QTY'].sum()
                if abs(pnet_total - factory_total) > 1e-9:
                    confirmation_list.append({
                        'PO#': po_num,
                        'PO-LINE#': line_num,
                        '사유': f'총수량 불일치(다운로드:{pnet_total}, 회신:{factory_total})'
                    })
                    continue

                upload_result_list.extend(self._map_factory_to_download(pnet_group, factory_group))

            self.result = pd.DataFrame(upload_result_list)
            self.confirmation_needed = pd.DataFrame(confirmation_list)
            self.change_summary = self._generate_change_summary(pnet_df, self.result)
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
                    'ETD(텍스트,YYYYMMDD)': download_row['ETD'],
                    'EX-F(텍스트,YYYYMMDD)': download_row['EX-F'],
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

        # Process smaller qty line first so they secure an ETD-group match;
        # larger lines absorb the remaining split burden.
        split_candidates.sort(key=lambda row: (row['_quantity'], row['_orig_index']))

        for download_row in split_candidates:
            self._assign_etd_grouped_split(download_row, factory_rows, result_rows)

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
            'ETD(텍스트,YYYYMMDD)': factory_row['ETD'],
            'EX-F(텍스트,YYYYMMDD)': factory_row['EX-F'],
            '내부노트': factory_row['내부노트'],
            'CPO#': download_row['CPO#'],
            'CPO-LINE#': download_row['CPO-LINE#'],
            'LINE SEQ': download_row['LINE SEQ']
        })

    def _assign_etd_grouped_split(
        self,
        download_row: Dict[str, Any],
        factory_rows: List[Dict[str, Any]],
        result_rows: List[Dict[str, Any]]
    ) -> None:
        """Assign download_row using ETD-group aggregation.
        Outputs ONE result row per ETD group used (not one per factory sub-row).
        Finds the minimum number of ETD groups needed to cover the required qty.
        """
        required_qty = download_row['_quantity']
        download_etd = str(download_row['ETD']).strip()

        # Build ETD groups from available factory rows
        groups: Dict[str, Dict] = {}
        for row in factory_rows:
            if row['remaining_qty'] > 0:
                etd = str(row['ETD']).strip()
                if etd not in groups:
                    groups[etd] = {
                        'etd': etd,
                        'rows': [],
                        'total_qty': 0.0,
                        'same_etd': etd == download_etd,
                    }
                groups[etd]['rows'].append(row)
                groups[etd]['total_qty'] += row['remaining_qty']

        if not groups:
            raise ValueError(
                f"Unable to fully map download row for PO#: {download_row['PO#']}, "
                f"PO-LINE#: {download_row['PO-LINE#']}, LINE SEQ: {download_row['LINE SEQ']}"
            )

        # Sort rows within each group largest-first
        for g in groups.values():
            g['rows'].sort(key=lambda r: -r['remaining_qty'])

        group_items = list(groups.values())
        best_combo = self._find_best_etd_group_combo(required_qty, group_items)

        if best_combo is None:
            raise ValueError(
                f"Unable to fully map download row for PO#: {download_row['PO#']}, "
                f"PO-LINE#: {download_row['PO-LINE#']}, LINE SEQ: {download_row['LINE SEQ']}"
            )

        # Sort selected groups: same-ETD first, then by descending total_qty
        selected = sorted(best_combo, key=lambda g: (not g['same_etd'], -g['total_qty']))

        remaining = required_qty
        for group in selected:
            if remaining <= 1e-9:
                break
            take = min(group['total_qty'], remaining)
            # Deduct from factory rows within the group
            to_deduct = take
            for row in group['rows']:
                if to_deduct <= 1e-9:
                    break
                qty = min(row['remaining_qty'], to_deduct)
                row['remaining_qty'] -= qty
                to_deduct -= qty
            # Primary row for Material/EX-F/내부노트
            primary = group['rows'][0]
            result_rows.append({
                'PO#': download_row['PO#'],
                'PO-LINE#': download_row['PO-LINE#'],
                'Material': primary['Material'],
                'CPO QTY': take,
                'ETD(텍스트,YYYYMMDD)': group['etd'],
                'EX-F(텍스트,YYYYMMDD)': primary['EX-F'],
                '내부노트': primary.get('내부노트'),
                'CPO#': download_row['CPO#'],
                'CPO-LINE#': download_row['CPO-LINE#'],
                'LINE SEQ': download_row['LINE SEQ']
            })
            remaining -= take

        if remaining > 1e-9:
            raise ValueError(
                f"Unable to fully map download row for PO#: {download_row['PO#']}, "
                f"PO-LINE#: {download_row['PO-LINE#']}, LINE SEQ: {download_row['LINE SEQ']}"
            )

    def _find_best_etd_group_combo(
        self,
        required_qty: float,
        group_items: List[Dict[str, Any]]
    ) -> Optional[List[Dict[str, Any]]]:
        """Find the minimum number of ETD groups whose combined total covers required_qty.
        Tie-break: prefer more same-ETD coverage, then larger primary group, then less surplus.
        """
        best_combo = None
        best_score = None

        for combo_size in range(1, len(group_items) + 1):
            for combo in combinations(group_items, combo_size):
                total = sum(g['total_qty'] for g in combo)
                if total + 1e-9 < required_qty:
                    continue
                same_etd_qty = sum(g['total_qty'] for g in combo if g['same_etd'])
                surplus = total - required_qty
                score = (
                    -combo_size,          # fewer ETD groups is better
                    same_etd_qty,         # more same-ETD coverage is better
                    max(g['total_qty'] for g in combo),  # larger primary group is better
                    -surplus,             # less waste is better
                )
                if best_score is None or score > best_score:
                    best_score = score
                    best_combo = combo

            if best_combo is not None:
                break  # found minimum combo_size; no need to try larger

        return list(best_combo) if best_combo else None

    def _generate_change_summary(self, original_pnet: pd.DataFrame, result_df: pd.DataFrame) -> pd.DataFrame:
        """Generate change summary comparing original and result."""
        summary = []
        if result_df.empty:
            return pd.DataFrame(summary)

        # Create key for comparison
        original_pnet['key'] = original_pnet['CPO#'] + '-' + original_pnet['CPO-LINE#'] + '-' + original_pnet['LINE SEQ']
        result_df['key'] = result_df['CPO#'] + '-' + result_df['CPO-LINE#'] + '-' + result_df['LINE SEQ']

        for key in result_df['key'].unique():
            orig_row = original_pnet[original_pnet['key'] == key]
            res_row = result_df[result_df['key'] == key]
            if orig_row.empty or res_row.empty:
                continue
            orig = orig_row.iloc[0]
            res = res_row.iloc[0]
            changes = []
            if orig['CPO QTY'] != res['CPO QTY']:
                changes.append(
                    f'수량: {self._format_quantity_for_change_summary(orig["CPO QTY"])}'
                    f' -> {self._format_quantity_for_change_summary(res["CPO QTY"])}'
                )

            orig_etd = self._normalize_date_for_change_summary(orig['ETD'])
            res_etd = self._normalize_date_for_change_summary(res['ETD(텍스트,YYYYMMDD)'])
            if orig_etd != res_etd:
                changes.append(f'ETD: {orig_etd} -> {res_etd}')

            if changes:
                summary.append({
                    'CPO#': res['CPO#'],
                    'CPO-LINE#': res['CPO-LINE#'],
                    'LINE SEQ': res['LINE SEQ'],
                    '변경종류': '; '.join(changes)
                })
        return pd.DataFrame(summary)

    def save_result(self, output_path: str) -> bool:
        """Save result to Excel file with multiple sheets, clearing existing sheets first."""
        try:
            if self.result is None:
                raise ValueError("No result to save. Run compare_and_generate first.")

            # 기존 파일이 있으면 열고 모든 시트 클리어
            try:
                wb = openpyxl.load_workbook(output_path)
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    ws.delete_rows(1, ws.max_row)
            except FileNotFoundError:
                wb = openpyxl.Workbook()
                # 기본 시트 제거
                wb.remove(wb.active)

            # 수동 업로드 sheet
            column_order = [
                'PO#', 'PO-LINE#', 'Material', 'CPO QTY', 'ETD(텍스트,YYYYMMDD)', 'EX-F(텍스트,YYYYMMDD)',
                '내부노트', 'CPO#', 'CPO-LINE#', 'LINE SEQ'
            ]
            ws_upload = wb.create_sheet('수동 업로드')
            df_upload = self.result[column_order]
            for r, row in enumerate(df_upload.itertuples(index=False), start=1):
                for c, value in enumerate(row, start=1):
                    ws_upload.cell(row=r, column=c, value=value)

            # 확인필요 sheet
            if self.confirmation_needed is not None and not self.confirmation_needed.empty:
                ws_confirm = wb.create_sheet('확인필요')
                for r, row in enumerate(self.confirmation_needed.itertuples(index=False), start=1):
                    for c, value in enumerate(row, start=1):
                        ws_confirm.cell(row=r, column=c, value=value)

            # 변경요약 sheet
            if self.change_summary is not None and not self.change_summary.empty:
                ws_summary = wb.create_sheet('변경요약')
                for r, row in enumerate(self.change_summary.itertuples(index=False), start=1):
                    for c, value in enumerate(row, start=1):
                        ws_summary.cell(row=r, column=c, value=value)

            wb.save(output_path)
            return True
        except Exception as e:
            print(f"Error saving result: {e}")
            return False
