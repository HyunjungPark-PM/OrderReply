#!/usr/bin/env python3
"""
p-net OrderReply Console Test Script
GUI 없이 콘솔에서 엑셀 처리 기능을 테스트합니다.
"""
from excel_processor import ExcelProcessor
import os
import pandas as pd

def test_excel_processing():
    print("=== p-net OrderReply 콘솔 테스트 ===")

    # 샘플 파일 존재 확인
    pnet_file = 'sample_pnet_download.xlsx'
    factory_file = 'sample_factory_reply.xlsx'
    output_file = 'test_result.xlsx'

    print(f"p-net 파일: {pnet_file} (존재: {os.path.exists(pnet_file)})")
    print(f"공장 파일: {factory_file} (존재: {os.path.exists(factory_file)})")

    # 엑셀 처리기 생성
    processor = ExcelProcessor()

    try:
        # 파일 읽기
        print("\n1. 파일 읽기...")
        if processor.read_pnet_download(pnet_file):
            print("✓ p-net 파일 읽기 성공")
            print(f"  - 행 수: {len(processor.pnet_download)}")
        else:
            print("✗ p-net 파일 읽기 실패")
            return

        if processor.read_factory_reply(factory_file):
            print("✓ 공장 파일 읽기 성공")
            print(f"  - 행 수: {len(processor.factory_reply)}")
        else:
            print("✗ 공장 파일 읽기 실패")
            return

        # 비교 및 생성
        print("\n2. 파일 비교 및 결과 생성...")
        if processor.compare_and_generate():
            print("✓ 파일 비교 및 생성 성공")
            print(f"  - 결과 행 수: {len(processor.result)}")
        else:
            print("✗ 파일 비교 및 생성 실패")
            return

        # 결과 저장
        print("\n3. 결과 저장...")
        if processor.save_result(output_file):
            print(f"✓ 결과 저장 성공: {output_file}")
        else:
            print("✗ 결과 저장 실패")
            return

        # 결과 확인
        print("\n4. 결과 확인...")
        if os.path.exists(output_file):
            result_df = pd.read_excel(output_file)
            print(f"✓ 결과 파일 생성됨: {len(result_df)} 행")
            print("\n결과 미리보기:")
            print(result_df.head().to_string(index=False))
        else:
            print("✗ 결과 파일이 생성되지 않음")

        print("\n=== 테스트 완료 ===")
        print("GUI가 작동하지 않는 경우 이 스크립트로 기능을 확인하세요.")

    except Exception as e:
        print(f"\n오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_excel_processing()