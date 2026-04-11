"""Debug script: trace mapping failures on actual files."""
import sys
import traceback
import pandas as pd
from itertools import combinations
from excel_processor import ExcelProcessor


def debug_group(po_num, line_num, download_rows_raw, factory_rows_raw):
    """Print detailed mapping trace for one PO+LINE group."""
    print(f"\n{'='*60}")
    print(f"PO#: {po_num}  LINE#: {line_num}")
    print(f"  Download rows ({len(download_rows_raw)}):")
    for r in download_rows_raw:
        print(f"    SEQ={r['LINE SEQ']:>5}  QTY={r['CPO QTY']:>8}  ETD={r['ETD']}")
    print(f"  Factory rows ({len(factory_rows_raw)}):")
    for r in factory_rows_raw:
        print(f"    QTY={r['CPO QTY']:>8}  ETD={r['ETD']}  EX-F={r['EX-F']}")

    dl_total = sum(float(r['CPO QTY'] or 0) for r in download_rows_raw)
    fa_total = sum(float(r['CPO QTY'] or 0) for r in factory_rows_raw)
    print(f"  Total QTY  download={dl_total}  factory={fa_total}  match={'OK' if abs(dl_total - fa_total) < 1e-9 else 'MISMATCH'}")

    # reproduce _map_factory_to_download step by step
    p = ExcelProcessor()
    dl_rows = []
    for i, r in enumerate(download_rows_raw):
        d = dict(r)
        d['_orig_index'] = i
        d['_quantity'] = float(r['CPO QTY'] or 0)
        dl_rows.append(d)

    fa_rows = []
    for i, r in enumerate(factory_rows_raw):
        qty = float(r['CPO QTY'] or 0)
        if qty > 0:
            fa_rows.append({**r, 'remaining_qty': qty, '_index': i})

    dl_rows.sort(key=lambda r: (-r['_quantity'], r['_orig_index']))

    result_rows = []
    split_candidates = []

    for dl in dl_rows:
        if dl['_quantity'] <= 0:
            result_rows.append({**dl})
            continue
        candidate = p._find_whole_assignment(dl, fa_rows)
        if candidate is not None:
            print(f"  [WHOLE] SEQ={dl['LINE SEQ']} QTY={dl['_quantity']} -> factory ETD={candidate['ETD']} remaining={candidate['remaining_qty']}")
            p._assign_factory_segment(dl, candidate, dl['_quantity'], result_rows)
        else:
            print(f"  [SPLIT candidate] SEQ={dl['LINE SEQ']} QTY={dl['_quantity']}")
            split_candidates.append(dl)

    split_candidates.sort(key=lambda r: (-r['_quantity'], r['_orig_index']))

    for dl in split_candidates:
        avail = [r for r in fa_rows if r['remaining_qty'] > 0]
        print(f"  [SPLIT attempt] SEQ={dl['LINE SEQ']} QTY={dl['_quantity']}")
        print(f"    Available factory remaining: {[(r['ETD'], r['remaining_qty']) for r in avail]}")
        grouped = p._find_split_group_combo(dl, fa_rows)
        if grouped is None:
            avail_total = sum(r['remaining_qty'] for r in fa_rows if r['remaining_qty'] > 0)
            print(f"    !! _find_split_group_combo returned None. avail_total={avail_total}")
        else:
            print(f"    group combo found: {[(r['ETD'], r['remaining_qty']) for g in grouped for r in g]}")
            try:
                p._assign_split_download(dl, fa_rows, result_rows)
                print(f"    Split assigned OK")
            except Exception as e:
                print(f"    !! Split FAILED: {e}")


def main():
    pnet_path = r"c:\Users\ryoum\OneDrive\바탕 화면\OrderReply\p-net 다운로드 파일.xlsx"
    factory_path = r"c:\Users\ryoum\OneDrive\바탕 화면\OrderReply\공장납기회신 파일.xlsx"

    p = ExcelProcessor()
    print("Loading files...")
    ok1 = p.read_pnet_download(pnet_path)
    ok2 = p.read_factory_reply(factory_path)
    if not ok1 or not ok2:
        print("File load failed")
        sys.exit(1)

    pnet_df = p.pnet_download.copy()
    factory_df = p.factory_reply.copy()

    pnet_df['PO#'] = pnet_df['PO#'].astype(str)
    pnet_df['PO-LINE#'] = pnet_df['PO-LINE#'].astype(str)
    factory_df['PO#'] = factory_df['PO#'].astype(str)
    factory_df['LINE#'] = factory_df['LINE#'].astype(str)

    error_groups = []
    ok_groups = 0

    for (po_num, line_num), pnet_group in pnet_df.groupby(['PO#', 'PO-LINE#'], sort=False):
        factory_group = factory_df[
            (factory_df['PO#'] == po_num) &
            (factory_df['LINE#'] == line_num)
        ].copy()

        if factory_group.empty:
            continue

        pnet_total = pnet_group['CPO QTY'].sum()
        factory_total = factory_group['CPO QTY'].sum()
        if abs(pnet_total - factory_total) > 1e-9:
            continue

        try:
            p._map_factory_to_download(pnet_group, factory_group)
            ok_groups += 1
        except Exception as e:
            error_groups.append((po_num, line_num, pnet_group, factory_group, str(e)))

    print(f"\nResult: {ok_groups} groups OK, {len(error_groups)} group(s) FAILED\n")

    # Show output summary for all groups
    print("\n--- Mapping output summary ---")
    p2 = ExcelProcessor()
    p2.read_pnet_download(pnet_path)
    p2.read_factory_reply(factory_path)
    p2.compare_and_generate()
    if p2.result is not None and not p2.result.empty:
        for _, row in p2.result.iterrows():
            orig = p2.pnet_download[
                (p2.pnet_download['CPO#'] == row['CPO#']) &
                (p2.pnet_download['CPO-LINE#'] == row['CPO-LINE#']) &
                (p2.pnet_download['LINE SEQ'] == row['LINE SEQ'])
            ]
            orig_qty = orig.iloc[0]['CPO QTY'] if not orig.empty else '?'
            orig_etd = orig.iloc[0]['ETD'] if not orig.empty else '?'
            out_qty  = row['CPO QTY']
            out_etd  = row['ETD(텍스트,YYYYMMDD)']
            qty_changed = str(orig_qty) != str(out_qty) and abs(float(orig_qty or 0) - float(out_qty or 0)) > 1e-9
            etd_changed = str(orig_etd).strip() != str(out_etd).strip()
            flag = []
            if qty_changed: flag.append("QTY_SPLIT")
            if etd_changed: flag.append("ETD_CHANGED")
            print(f"  SEQ={row['LINE SEQ']:>5}  orig_qty={orig_qty}  out_qty={out_qty}  "
                  f"orig_etd={orig_etd}  out_etd={out_etd}  {'  '.join(flag) or 'OK'}")

    for (po_num, line_num, pg, fg, err) in error_groups:
        fa_records_fixed = fg.to_dict('records')
        debug_group(po_num, line_num, pg.to_dict('records'), fa_records_fixed)
        print(f"\n  >> Error message: {err}")


if __name__ == '__main__':
    main()
