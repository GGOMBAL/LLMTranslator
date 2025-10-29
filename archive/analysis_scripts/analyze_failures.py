#!/usr/bin/env python3
"""
번역 실패 원인 분석 스크립트
"""

import json
import pandas as pd
from collections import defaultdict, Counter
import re

def load_translation_data(json_path):
    """JSON 파일에서 번역 데이터 로드"""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def analyze_failure_patterns(data):
    """실패 패턴 분석"""

    failed_pages = []
    failure_types = defaultdict(list)

    for page in data['pages']:
        translated = page['translated_text']

        # 실패 케이스 감지
        is_failed = False
        failure_reason = None

        if '[Translation failed' in translated:
            is_failed = True
            # 실패 메시지 추출
            match = re.search(r'\[Translation failed.*?\]', translated)
            if match:
                failure_reason = match.group(0)

        elif '[TOC -' in translated:
            is_failed = True
            failure_reason = '[TOC Translation Failed]'

        if is_failed:
            failed_pages.append({
                'page_number': page['page_number'],
                'original_length': page['original_char_count'],
                'failure_message': failure_reason,
                'original_preview': page['original_text'][:200],
                'translation_time': page['translation_time']
            })

            # 실패 유형 분류
            if 'after 2 attempts' in translated:
                failure_types['재시도 후 실패'].append(page['page_number'])
            elif 'sequence item' in translated or 'NoneType' in translated:
                failure_types['타입 에러'].append(page['page_number'])
            elif 'TOC' in translated:
                failure_types['목차(TOC) 처리 실패'].append(page['page_number'])
            else:
                failure_types['기타'].append(page['page_number'])

    return failed_pages, failure_types

def categorize_content_types(failed_pages):
    """실패한 페이지의 콘텐츠 유형 분류"""

    content_patterns = {
        '목차/인덱스': [r'目\s*录', r'序号.*版本.*变更内容', r'^\s*-\s*\d+\s*-'],
        '표/테이블': [r'表\d+', r'┃', r'│', r'工况', r'前提条件.*测试工况.*目标要求'],
        '기술 사양': [r'ATS模式', r'IBC', r'功能要求', r'状态', r'TBD', r'条件'],
        '긴 리스트': [r'^\d+\..*\n\d+\..*\n\d+\.', r'[a-z]\.\s.*\n[a-z]\.\s'],
    }

    categorized = defaultdict(list)

    for page in failed_pages:
        original = page['original_preview']
        page_num = page['page_number']

        matched = False
        for category, patterns in content_patterns.items():
            for pattern in patterns:
                if re.search(pattern, original):
                    categorized[category].append(page_num)
                    matched = True
                    break
            if matched:
                break

        if not matched:
            categorized['일반 텍스트'].append(page_num)

    return categorized

def analyze_length_correlation(failed_pages):
    """실패와 원문 길이의 상관관계 분석"""

    lengths = [p['original_length'] for p in failed_pages]

    length_analysis = {
        '평균 길이': sum(lengths) / len(lengths) if lengths else 0,
        '최소 길이': min(lengths) if lengths else 0,
        '최대 길이': max(lengths) if lengths else 0,
        '중앙값': sorted(lengths)[len(lengths)//2] if lengths else 0,
    }

    # 길이별 분포
    length_distribution = {
        '매우 짧음 (<500자)': sum(1 for l in lengths if l < 500),
        '짧음 (500-1000자)': sum(1 for l in lengths if 500 <= l < 1000),
        '중간 (1000-2000자)': sum(1 for l in lengths if 1000 <= l < 2000),
        '긴편 (2000-3000자)': sum(1 for l in lengths if 2000 <= l < 3000),
        '매우 긺 (3000자+)': sum(1 for l in lengths if l >= 3000),
    }

    return length_analysis, length_distribution

def generate_analysis_report(data):
    """종합 분석 보고서 생성"""

    print("="*80)
    print("📊 번역 실패 원인 분석 보고서")
    print("="*80)

    # 1. 기본 통계
    total_pages = data['total_pages_processed']
    success_pages = data['successful_translations']
    failed_count = total_pages - success_pages

    print(f"\n📈 기본 통계")
    print(f"  - 총 페이지: {total_pages}")
    print(f"  - 성공: {success_pages} ({success_pages/total_pages*100:.1f}%)")
    print(f"  - 실패: {failed_count} ({failed_count/total_pages*100:.1f}%)")

    # 2. 실패 패턴 분석
    failed_pages, failure_types = analyze_failure_patterns(data)

    print(f"\n🔍 실패 유형 분류")
    for failure_type, pages in sorted(failure_types.items(), key=lambda x: len(x[1]), reverse=True):
        print(f"  - {failure_type}: {len(pages)}건")
        print(f"    페이지: {', '.join(map(str, sorted(pages)[:10]))}" +
              (f" ... 외 {len(pages)-10}건" if len(pages) > 10 else ""))

    # 3. 콘텐츠 유형 분석
    content_categories = categorize_content_types(failed_pages)

    print(f"\n📄 실패한 콘텐츠 유형")
    for content_type, pages in sorted(content_categories.items(), key=lambda x: len(x[1]), reverse=True):
        print(f"  - {content_type}: {len(pages)}건")
        print(f"    페이지: {', '.join(map(str, sorted(pages)[:10]))}" +
              (f" ... 외 {len(pages)-10}건" if len(pages) > 10 else ""))

    # 4. 길이 상관관계 분석
    length_analysis, length_distribution = analyze_length_correlation(failed_pages)

    print(f"\n📏 실패 페이지 길이 분석")
    for metric, value in length_analysis.items():
        print(f"  - {metric}: {value:.0f}자")

    print(f"\n📊 길이별 실패 분포")
    for range_name, count in length_distribution.items():
        percentage = (count / len(failed_pages) * 100) if failed_pages else 0
        bar = "█" * int(percentage / 5)
        print(f"  - {range_name:20s}: {count:3d}건 ({percentage:5.1f}%) {bar}")

    # 5. 상세 실패 케이스 샘플
    print(f"\n🔬 상세 실패 케이스 (샘플 5건)")
    for i, page in enumerate(failed_pages[:5], 1):
        print(f"\n  [{i}] 페이지 {page['page_number']}")
        print(f"      원문 길이: {page['original_length']}자")
        print(f"      처리 시간: {page['translation_time']:.2f}초")
        print(f"      실패 메시지: {page['failure_message']}")
        print(f"      원문 미리보기: {page['original_preview'][:150]}...")

    # 6. 주요 실패 원인 요약
    print(f"\n💡 주요 실패 원인 분석")

    # 원인별 분석
    reasons = []

    if '목차(TOC) 처리 실패' in failure_types and len(failure_types['목차(TOC) 처리 실패']) > 0:
        reasons.append({
            'reason': '목차(TOC) 구조 처리 오류',
            'description': '복잡한 목차 구조에서 NoneType 에러 발생',
            'affected': len(failure_types['목차(TOC) 처리 실패']),
            'solution': '목차 처리 로직 개선 필요 (None 값 핸들링)'
        })

    if '재시도 후 실패' in failure_types and len(failure_types['재시도 후 실패']) > 0:
        reasons.append({
            'reason': '번역 API 2회 시도 후 실패',
            'description': 'API 호출이 2번 모두 실패하여 번역 불가',
            'affected': len(failure_types['재시도 후 실패']),
            'solution': '재시도 횟수 증가, 청크 크기 조정, 타임아웃 증가'
        })

    if '표/테이블' in content_categories and len(content_categories['표/테이블']) > 5:
        reasons.append({
            'reason': '복잡한 표/테이블 구조',
            'description': '표 형태의 데이터가 번역 중 구조 손실',
            'affected': len(content_categories['표/테이블']),
            'solution': '표 전용 처리 로직 추가, 구조 보존 알고리즘 적용'
        })

    # 긴 텍스트 분석
    long_texts = [p for p in failed_pages if p['original_length'] > 1000]
    if len(long_texts) > 10:
        reasons.append({
            'reason': '긴 텍스트 처리 실패',
            'description': f'1000자 이상의 긴 텍스트가 {len(long_texts)}건 실패',
            'affected': len(long_texts),
            'solution': '청크 분할 전략 개선, 컨텍스트 유지 방법 개선'
        })

    for i, reason_info in enumerate(reasons, 1):
        print(f"\n  {i}. {reason_info['reason']}")
        print(f"     설명: {reason_info['description']}")
        print(f"     영향: {reason_info['affected']}건")
        print(f"     해결방안: {reason_info['solution']}")

    # 7. 권장사항
    print(f"\n✅ 개선 권장사항")
    print(f"  1. 목차/TOC 처리 로직에서 None 값 체크 강화")
    print(f"  2. 재시도 메커니즘 개선 (지수 백오프, 재시도 횟수 증가)")
    print(f"  3. 복잡한 표 구조 전용 처리 파이프라인 추가")
    print(f"  4. 긴 텍스트는 더 작은 청크로 분할")
    print(f"  5. API 타임아웃 설정 증가")
    print(f"  6. 실패 시 폴백(fallback) 번역 서비스 사용")

    print("\n" + "="*80)

    return failed_pages, failure_types, content_categories

def main():
    """메인 실행 함수"""

    json_path = '/Users/kimjinmyung/Desktop/GIT/TranslatewithRag/final_translation_results.json'

    data = load_translation_data(json_path)
    failed_pages, failure_types, content_categories = generate_analysis_report(data)

    # 추가 CSV 출력 (실패 상세 내역)
    if failed_pages:
        df_failures = pd.DataFrame(failed_pages)
        csv_path = '/Users/kimjinmyung/Desktop/GIT/TranslatewithRag/실패_상세분석.csv'
        df_failures.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"\n📁 상세 실패 분석 CSV 저장: {csv_path}")

if __name__ == '__main__':
    main()
