# -*- coding: utf-8 -*-

import os
import sys
import datetime
import traceback
import xlrd as ModXlrd

Book = ModXlrd.book
Sheet = ModXlrd.sheet

HEADER_ROW = 4

class CXlsReader(object):
	def __init__(self):
		self.book = None
		self.sheets = {}

	def IsValidSheet(self, sheet):
		return sheet.nrows > HEADER_ROW and sheet.ncols > 1

	def GetRowCnt(self, sheet):
		colVal = sheet.col_values(0)
		for i in range(HEADER_ROW, sheet.nrows):
			if "" == colVal[i]:
				return i
		return sheet.nrows

	def GetColCnt(self, sheet):
		rowVal = sheet.row_values(1) # 第二行，types
		for i in range(sheet.ncols):
			if "" == rowVal[i]:
				return i
		return sheet.ncols

	def ConvertOneSheet(self, fp, sheet):
		# headers, types, defaults, logicNames
		nRow = self.GetRowCnt(sheet)
		nCol = self.GetColCnt(sheet)
		headers = [v for v in sheet.row_values(0)[:nCol]]
		types = [v for v in sheet.row_values(1)[:nCol]]
		defaults = [str(v) for v in sheet.row_values(2)[:nCol]]
		logicNames = [v for v in sheet.row_values(3)[:nCol]]
		export_cols = ["false" if i == 0 else "true" for i in range(nCol)]
		lines = []
		for i in range(HEADER_ROW, nRow):
			rowData = [str(v) for v in sheet.row_values(i)[:nCol]]
			for j in range(len(rowData)):
				if "" == rowData[j]:
					rowData[j] = defaults[j]
			lines.append(rowData)

		fp.write('["%s"] = {\n' % (sheet.name.lower()))
		fp.write('\tNAME = "%s",\n' % (sheet.name.lower()))
		fp.write("\tROW_CNT = %d,\n\tCOL_CNT = %s,\n" % (nRow, nCol))
		self.WriteStrList(fp, "HEADER", headers)
		self.WriteStrList(fp, "TYPE", types)
		self.WriteStrList(fp, "LOGICNAME", logicNames)
		self.WriteStrList(fp, "DEFAULT", defaults)
		self.WriteStrList(fp, "EXPORT_COLS", export_cols, False)
		fp.write("\tDATA = {\n")
		for rowData in lines:
			self.WriteCfgDatas(fp, rowData)
		fp.write("\t},\n}, -- End of sheet %s \n" % (sheet.name.lower()))

	def OpenXLS(self, xlsFullPath, targetFile):
		if not os.path.exists(xlsFullPath):
			print("Xls file NOT exists! file: %s" % (xlsFullPath))
			sys.exit(1)
		self.book = ModXlrd.open_workbook(xlsFullPath)
		# 写 lua 文件
		fp = open(targetFile, "w", encoding = "utf8")
		_, xlsFileName = os.path.split(xlsFullPath)
		xlsName, _ = os.path.splitext(xlsFileName)
		xlsFileName = xlsFileName.lower()
		xlsName = xlsName.lower()
		fp.write("-- File: %s\n" % (os.path.normpath(targetFile)))
		fp.write("-- Desc: excel 文件 %s 导出的中间数据\n" % (xlsFileName))
		fp.write("-- Date: %s\n\n" % (datetime.datetime.now()))
		fp.write("%s = {\n" % (xlsName))
		fp.write('XLS_FILE = "%s",\nXLS_NAME = "%s",\nHEADER_ROW = %d,\nSHEETS = {\n' % (xlsFileName, xlsName, HEADER_ROW))
		validSheet = 0
		for sheet in self.book.sheets():
			if not self.IsValidSheet(sheet): # 低于最少行数
				continue
			validSheet += 1
			try:
				self.ConvertOneSheet(fp, sheet)
			except:
				seperator = "\n%s" % ("*" * 80)
				print(seperator)
				print("\n\t\t\t!!!Error occurs while converting!!!\n\t\t\texcel=%s, sheet=%s\nError=%s\n\n" % (xlsFileName.upper(), sheet.name.upper(), traceback.format_exc()))
				print(seperator)
				sys.exit(1)
		fp.write("}, -- End of xls %s, valid sheet amount %d \n-- End of file" % (xlsFileName, validSheet))
		fp.write("\n}\n")
		fp.close()
		return True

	def WriteStrList(self, fp, name, strList, bQuote = True):
		fp.write("\t%s = {" % (name))
		for val in strList:
			if bQuote:
				fp.write('"%s", ' % (val))
			else:
				fp.write('%s, ' % (val))
		fp.write("},\n")

	def escape(self, s):
		return s.replace('"', r'\"').replace('\n', r'\n')

	def WriteCfgDatas(self, fp, rowData):
		fp.write("\t\t")
		fp.write('["%s"]' % self.escape(rowData[0]))
		fp.write(" = {")
		for i in range(0, len(rowData)):
			fp.write('"%s", ' % self.escape(rowData[i]))
		fp.write("},\n")

if __name__ == "__main__":
	xls = CXlsReader()
	xls.OpenXLS(sys.argv[1], sys.argv[2])
