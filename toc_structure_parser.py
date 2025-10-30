#!/usr/bin/env python3
"""
목차 구조 파서 (TOC Structure Parser)

기능:
- 3.1.1, 2.3, 1.2.3.4 같은 계층 구조 인식
- 페이지별 섹션 매핑
- 계층 구조 분석
"""

import re
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

@dataclass
class TOCItem:
    """목차 항목"""
    number: str          # "3.1.1"
    title: str           # "시스템 개요"
    level: int           # 3 (점의 개수 + 1)
    page: Optional[int]  # 페이지 번호
    parent: Optional[str] # 부모 번호 "3.1"

class TOCStructureParser:
    """목차 구조 파서"""

    def __init__(self):
        self.toc_items: List[TOCItem] = []
        self.section_patterns = [
            # 3.1.1, 2.3.4.5 등
            r'(\d+(?:\.\d+){1,4})\s+(.+?)(?:\s+\.{2,}|\s+(?=\d+$)|$)',
            # 1.2.3 Title ... 45 (페이지 번호 포함)
            r'(\d+(?:\.\d+){1,4})\s+(.+?)\s+\.{2,}\s*(\d+)',
            # 단순 번호 + 제목
            r'(\d+(?:\.\d+){1,4})\s+([^\d]+?)(?:\s+\d+)?$',
        ]

    def parse_toc_text(self, toc_text: str) -> List[TOCItem]:
        """
        목차 텍스트 파싱

        Args:
            toc_text: 목차 페이지의 텍스트

        Returns:
            TOCItem 리스트
        """
        lines = toc_text.split('\n')
        items = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 패턴 매칭 시도
            for pattern in self.section_patterns:
                match = re.search(pattern, line)
                if match:
                    groups = match.groups()
                    number = groups[0]
                    title = groups[1].strip()
                    page = int(groups[2]) if len(groups) > 2 and groups[2] else None

                    # 레벨 계산 (점의 개수 + 1)
                    level = number.count('.') + 1

                    # 부모 번호 계산
                    parent = '.'.join(number.split('.')[:-1]) if '.' in number else None

                    item = TOCItem(
                        number=number,
                        title=title,
                        level=level,
                        page=page,
                        parent=parent
                    )
                    items.append(item)
                    break

        self.toc_items = items
        return items

    def extract_section_from_text(self, text: str, debug: bool = False) -> Optional[str]:
        """
        페이지 텍스트에서 섹션 번호 추출

        Args:
            text: 페이지 텍스트
            debug: 디버그 모드

        Returns:
            섹션 번호 (예: "3.1.1") 또는 None
        """
        # 텍스트 시작 부분에서 섹션 번호 찾기 (범위 확대: 500 → 2000자)
        search_text = text[:2000]

        patterns = [
            # 기본 패턴
            r'^(\d+(?:\.\d+){0,4})\s+[^\d\n]',  # 시작 부분: "3.1.1 제목"
            r'\n(\d+(?:\.\d+){0,4})\s+[^\d\n]',  # 줄 시작: "\n3.1.1 제목"

            # 중국어 패턴
            r'第?\s*(\d+(?:\.\d+){0,4})\s*章',  # "第3.1章"
            r'(\d+(?:\.\d+){0,4})\s*[、，]',  # "3.1、内容"

            # 일반적인 패턴
            r'(?:^|\n)(\d+(?:\.\d+){0,4})\s*[\u4e00-\u9fff]',  # 섹션 + 중국어
            r'(?:^|\n)(\d+(?:\.\d+){0,4})\s+[A-Z]',  # 섹션 + 영어 대문자

            # 느슨한 패턴
            r'(?:^|\n)(\d+(?:\.\d+){1,4})(?:\s|$)',  # 섹션 번호만
        ]

        for pattern in patterns:
            match = re.search(pattern, search_text)
            if match:
                section = match.group(1)
                # 유효성 검사: 너무 긴 숫자는 제외
                if len(section) <= 10 and section.count('.') <= 4:
                    if debug:
                        print(f"      섹션 감지: {section} (패턴: {pattern[:30]}...)")
                    return section

        return None

    def get_section_info(self, section_number: str) -> Optional[TOCItem]:
        """
        섹션 번호로 목차 정보 가져오기

        Args:
            section_number: 섹션 번호 (예: "3.1.1")

        Returns:
            TOCItem 또는 None
        """
        for item in self.toc_items:
            if item.number == section_number:
                return item
        return None

    def build_hierarchy(self) -> Dict:
        """
        계층 구조 딕셔너리 생성

        Returns:
            {
                "1": {
                    "item": TOCItem,
                    "children": {
                        "1.1": {...},
                        "1.2": {...}
                    }
                }
            }
        """
        hierarchy = {}

        for item in self.toc_items:
            # 최상위 레벨
            if item.level == 1:
                hierarchy[item.number] = {
                    "item": item,
                    "children": {}
                }
            else:
                # 부모 찾기
                parent_dict = hierarchy
                parts = item.number.split('.')

                for i in range(len(parts) - 1):
                    parent_num = '.'.join(parts[:i+1])
                    if parent_num in parent_dict:
                        if "children" not in parent_dict[parent_num]:
                            parent_dict[parent_num]["children"] = {}
                        parent_dict = parent_dict[parent_num]["children"]

                parent_dict[item.number] = {
                    "item": item,
                    "children": {}
                }

        return hierarchy

    def map_pages_to_sections(self, pages_data: List[Dict], debug: bool = False) -> Dict[int, str]:
        """
        페이지를 섹션에 매핑

        Args:
            pages_data: 페이지 데이터 리스트
            debug: 디버그 모드

        Returns:
            {page_number: section_number}
        """
        page_to_section = {}
        detected_count = 0

        if debug:
            print(f"\n   🔍 섹션 감지 시작 (총 {len(pages_data)} 페이지)")

        for page in pages_data:
            page_num = page['page_number']
            text = page.get('original_text', '')

            # 텍스트에서 섹션 번호 추출
            section = self.extract_section_from_text(text, debug=debug)
            if section:
                page_to_section[page_num] = section
                detected_count += 1
                if debug:
                    print(f"   ✅ 페이지 {page_num}: 섹션 {section}")

        if debug:
            print(f"   📊 직접 감지: {detected_count}개 페이지")

        # 섹션이 없는 페이지는 이전 섹션을 계속 사용
        current_section = None
        filled_count = 0

        for page in sorted(pages_data, key=lambda x: x['page_number']):
            page_num = page['page_number']
            if page_num in page_to_section:
                current_section = page_to_section[page_num]
            elif current_section:
                page_to_section[page_num] = current_section
                filled_count += 1

        if debug:
            print(f"   📊 이전 섹션 사용: {filled_count}개 페이지")
            print(f"   📊 총 매핑: {len(page_to_section)}개 페이지\n")

        return page_to_section

    def get_hierarchy_path(self, section_number: str) -> List[str]:
        """
        섹션의 계층 경로 가져오기

        Args:
            section_number: "3.1.2"

        Returns:
            ["3", "3.1", "3.1.2"]
        """
        parts = section_number.split('.')
        path = []
        for i in range(len(parts)):
            path.append('.'.join(parts[:i+1]))
        return path

    def format_hierarchy_text(self, section_number: str, title: str = "") -> str:
        """
        계층 구조를 텍스트로 포맷팅

        Args:
            section_number: "3.1.2"
            title: 섹션 제목

        Returns:
            "    3.1.2 Title" (들여쓰기 포함)
        """
        level = section_number.count('.') + 1
        indent = "  " * (level - 1)
        return f"{indent}{section_number} {title}".strip()


def test_parser():
    """테스트 함수"""
    parser = TOCStructureParser()

    # 테스트 목차
    toc_text = """
    1 系统概述 ........................... 1
    1.1 项目背景 ......................... 2
    1.2 系统架构 ......................... 3
    2 功能需求 ........................... 5
    2.1 用户管理 ......................... 6
    2.1.1 用户注册 ....................... 7
    2.1.2 用户登录 ....................... 8
    2.2 数据处理 ......................... 10
    3 技术规范 ........................... 15
    """

    items = parser.parse_toc_text(toc_text)

    print("=" * 80)
    print("목차 구조 파싱 결과")
    print("=" * 80)

    for item in items:
        indent = "  " * (item.level - 1)
        print(f"{indent}{item.number} {item.title} (Level {item.level}, Page {item.page})")

    print("\n계층 구조:")
    hierarchy = parser.build_hierarchy()
    print(hierarchy)

    # 섹션 추출 테스트
    test_texts = [
        "3.1.1 系统概述\n这是内容...",
        "第2.3章 技术规范\n详细说明...",
        "1.2 用户管理系统\n功能描述..."
    ]

    print("\n섹션 번호 추출 테스트:")
    for text in test_texts:
        section = parser.extract_section_from_text(text)
        print(f"  '{text[:30]}...' -> {section}")


if __name__ == "__main__":
    test_parser()
