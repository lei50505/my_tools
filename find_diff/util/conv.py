#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''


def to_float(val):
    '''doc'''
    # pylint:disable=broad-except
    try:
        return float(val)
    except Exception:
        return None

def to_str(val):
    '''doc'''
    # pylint:disable=broad-except
    if val is None:
        return None
    try:
        val = str(val)
        if val.strip() == "":
            return None
        return val.strip()
    except Exception:
        return None
