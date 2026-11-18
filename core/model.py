"""공통 중간 모델: PPT/HTML 어느 쪽에서 읽어도 이 구조로 변환한다."""
from dataclasses import dataclass, field
from typing import List, Tuple


@dataclass
class Block:
    kind: str                      # 'para' | 'list' | 'table' | 'image'
    text: str = ""                 # para
    items: List[Tuple[int, str]] = field(default_factory=list)  # list: (level, text)
    ordered: bool = False          # list
    rows: List[List[str]] = field(default_factory=list)         # table
    header: bool = True            # table 첫 행이 머리글인가
    data: bytes = b""              # image
    ext: str = "png"               # image
    alt: str = ""


@dataclass
class Section:
    title: str = ""
    blocks: List[Block] = field(default_factory=list)


@dataclass
class Document:
    title: str = ""
    sections: List[Section] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
