#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''

import time
from functools import wraps

def fn_time(func):
    '''doc'''
    @wraps(func)
    def function_time(*args, **kw):
        '''doc'''
        start_time = time.time()
        result = func(*args, **kw)
        end_time = time.time()
        print("[%s]:[%s seconds]" \
            % (func.__name__, str(end_time - start_time)))
        return result
    return function_time
