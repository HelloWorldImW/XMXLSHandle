#! /usr/bin/env python
# _*_ coding:utf-8 _*_

'ExlModel数据处理'

__author__ = 'DarrenW'

import XMExlHandle

def goFly(self):
    pass


if __name__ == '__main__':
    path = './'
    b = raw_input('周报是否在当前目录下?(y/n)')
    if b == 'y':
        pass
    else:
        path = raw_input('请输入周报路径:')
    a = XMExlHandle.ExlReadHandle()
    m =  a.findExlFile(path)

    XMExlHandle.ExlWriteHandle(m)