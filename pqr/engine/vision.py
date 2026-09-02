# -*- coding: utf-8 -*-
"""손글씨 안정성시험일지 판독 (Claude 비전). ANTHROPIC_API_KEY 가 있을 때만 켜진다."""
import os


def available():
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def hook():
    from .vision_claude import read_stability_into
    return read_stability_into
