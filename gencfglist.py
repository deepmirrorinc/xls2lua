# -*- coding: utf-8 -*-

import os
import sys

if __name__ == "__main__":
    path = os.path.normpath(sys.argv[1])
    ext = sys.argv[2]
    target = sys.argv[3].lower()
    cfgList = []
    for fileItem in os.listdir(path):
        if os.path.isfile("%s/%s" % (path, fileItem)):
            if fileItem.lower() != target:
                cfgList.append(fileItem.lower())
    fp = open("%s/%s" % (path, target), "w", encoding = "utf8")
    fp.write("-- File: %s\n" % (target))
    fp.write("-- Desc: config file list\n\n")
    fp.write("---------------一下代码由工具生成，请不要修改---------------\n\n\n")
    fp.write('__GLOBALCONFIGTABLE__ = {}\n')
    for fileItem in cfgList:
        if "datamgr" in fileItem:
            continue
        fp.write('require("%s")\n' % fileItem[:-4])
    fp.close()
