import pandas as pd
from itertools import combinations
from typing import Any, List, Dict, Optional
from datetime import datetime


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
        
        # Handle string input
        if isinstance(date_val, str):
            date_val = date_val.strip()
            # Already in YYYYMMDD format
            if len(date_val) == 8 and date_val.isdigit():
                return date_val
            # Try parsing as datetime
            try:
                parsed = pd.to_datetime(date_val)
                return parsed.strftime('%Y%m%d')
            except:
                return date_val
        
        # Handle numeric input (Excel serial or YYYYMMDD as int)
        try:
            num_val = int(float(date_val))
            # If number is in YYYYMMDD format (8 digits, between 19000101 and 29991231)
            if 19000101 <= num_val <= 29991231 and len(str(num_val)) == 8:
                return str(num_val)
            # Otherwise treat as Excel serial number
            parsed = pd.to_datetime(num_val, unit='D', origin=pd.Timestamp("1899-12-30"))
            return parsed.strftime('%Y%m%d')
        except:
            return str(date_val)

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

                # Try global no-split assignment
                assignment = self._try_global_no_split(pnet_group, factory_group)
                if assignment:
                    upload_result_list.extend(assignment)
                else:
                    # Try split assignment
                    split_assignment = self._try_split_assignment(pnet_group, factory_group)
                    if split_assignment:
                        upload_result_list.extend(split_assignment)
                    else:
                        confirmation_list.append({
                            'PO#': po_num,
                            'PO-LINE#': line_num,
                            '사유': '공장납기회신 파일의 수량 구조상 기존 LINE SEQ조합으로 배정 불가 / 최소분할원칙위반 (단일 LINE SEQ 2~3분할로도 회신 수량 구성 불가)'
                        })

            self.result = pd.DataFrame(upload_result_list)
            self.confirmation_needed = pd.DataFrame(confirmation_list)
            self.change_summary = self._generate_change_summary(pnet_df, self.result)
            return True
        except Exception as e:
            print(f"Error comparing files: {e}")
            return False

    def _try_global_no_split(self, pnet_group: pd.DataFrame, factory_group: pd.DataFrame) -> Optional[List[Dict[str, Any]]]:
        """Try global no-split assignment."""
        # Group factory by ETD
        factory_by_etd = factory_group.groupby('ETD')
        etd_requirements = {etd: group['CPO QTY'].sum() for etd, group in factory_by_etd}

        # Find subsets of pnet LINE SEQ that match ETD requirements
        pnet_rows = pnet_group.to_dict('records')
        for etd, required_qty in etd_requirements.items():
            matching_rows = [row for row in pnet_rows if row['ETD'] == etd and row['CPO QTY'] == required_qty]
            if matching_rows:
                # Assign
                result = []
                for row in matching_rows:
                    factory_row = factory_group[factory_group['ETD'] == etd].iloc[0]  # Take first
                    result.append({
                        'PO#': row['PO#'],
                        'PO-LINE#': row['PO-LINE#'],
                        'Material': row['Material'],
                        'CPO QTY': row['CPO QTY'],
                        'ETD(텍스트,YYYYMMDD)': factory_row['ETD'],
                        'EX-F(텍스트,YYYYMMDD)': factory_row['EX-F'],
                        '내부노트': factory_row['내부노트'],
                        'CPO#': row['CPO#'],
                        'CPO-LINE#': row['CPO-LINE#'],
                        'LINE SEQ': row['LINE SEQ']
                    })
                # Remove assigned
                pnet_rows = [row for row in pnet_rows if row not in matching_rows]
                if not pnet_rows:
                    return result
        return None
    def _try_split_assignment(self, pnet_group: pd.DataFrame, factory_group: pd.DataFrame) -> Optional[List[Dict[str, Any]]]:
        """Try split assignment with conditions."""
        # Simplified split logic: allow splitting into 2-3 parts
        pnet_rows = pnet_group.to_dict('records')
        factory_rows = factory_group.to_dict('records')

        # Check if split is possible
        pnet_total = sum(row['CPO QTY'] for row in pnet_rows)
        factory_total = sum(row['CPO QTY'] for row in factory_rows)
        if pnet_total < factory_total:
            return None

        # Try to assign with splitting
        result = []
        factory_idx = 0
        for pnet_row in pnet_rows:
            qty_needed = pnet_row['CPO QTY']
            while qty_needed > 0 and factory_idx < len(factory_rows):
                factory_row = factory_rows[factory_idx]
                assign_qty = min(qty_needed, factory_row['CPO QTY'])
                result.append({
                    'PO#': pnet_row['PO#'],
                    'PO-LINE#': pnet_row['PO-LINE#'],
                    'Material': pnet_row['Material'],
                    'CPO QTY': assign_qty,
                    'ETD(텍스트,YYYYMMDD)': factory_row['ETD'],
                    'EX-F(텍스트,YYYYMMDD)': factory_row['EX-F'],
                    '내부노트': factory_row['내부노트'],
                    'CPO#': pnet_row['CPO#'],
                    'CPO-LINE#': pnet_row['CPO-LINE#'],
                    'LINE SEQ': pnet_row['LINE SEQ']
                })
                qty_needed -= assign_qty
                factory_row['CPO QTY'] -= assign_qty
                if factory_row['CPO QTY'] <= 0:
                    factory_idx += 1
            if qty_needed > 0:
                return None  # Cannot assign fully
        return result

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
                changes.append(f'수량: {orig["CPO QTY"]} -> {res["CPO QTY"]}')
            if orig['ETD'] != res['ETD(텍스트,YYYYMMDD)']:
                changes.append(f'ETD: {orig["ETD"]} -> {res["ETD(텍스트,YYYYMMDD)"]}')
            if changes:
                summary.append({
                    'CPO#': res['CPO#'],
                    'CPO-LINE#': res['CPO-LINE#'],
                    'LINE SEQ': res['LINE SEQ'],
                    '변경종류': '; '.join(changes)
                })
        return pd.DataFrame(summary)

    def save_result(self, output_path: str) -> bool:
        """Save result to Excel file with multiple sheets."""
        try:
            if self.result is None:
                raise ValueError("No result to save. Run compare_and_generate first.")

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 수동 업로드 sheet
                column_order = [
                    'PO#', 'PO-LINE#', 'Material', 'CPO QTY', 'ETD(텍스트,YYYYMMDD)', 'EX-F(텍스트,YYYYMMDD)',
                    '내부노트', 'CPO#', 'CPO-LINE#', 'LINE SEQ'
                ]
                self.result[column_order].to_excel(writer, sheet_name='수동 업로드', index=False)

                # 확인필요 sheet
                if self.confirmation_needed is not None and not self.confirmation_needed.empty:
                    self.confirmation_needed.to_excel(writer, sheet_name='확인필요', index=False)

                # 변경요약 sheet
                if self.change_summary is not None and not self.change_summary.empty:
                    self.change_summary.to_excel(writer, sheet_name='변경요약', index=False)

            return True
        except Exception as e:
            print(f"Error saving result: {e}")
            return False
