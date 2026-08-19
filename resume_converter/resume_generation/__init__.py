"""
标准简历Word生成公开接口。
"""

from .chery_generator import CheryResumeGenerator
from .generator import ResumeGenerator, YouzuResumeGenerator

__all__ = [
    "ResumeGenerator",
    "YouzuResumeGenerator",
    "CheryResumeGenerator",
]
