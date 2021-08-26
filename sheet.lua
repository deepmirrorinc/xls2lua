

CSheet = Class("CSheet")

function CSheet:CSheet(xls, sheetObj)
    self.xls = xls
    self.sheet = sheetObj
    self.curCol = 1
end

function CSheet:GetXlsFileName()
    return self.xls.XLS_FILE
end

function CSheet:GetXlsName()
    return self.xls.XLS_NAME
end

function CSheet:GetName()
    return self.sheet.NAME
end

function CSheet:GetFirstColName()
    return self.sheet.LOGICNAME[1]
end

function CSheet:GetLogicNames()
    return self.sheet.LOGICNAME
end

function CSheet:GetColTypes()
    return self.sheet.TYPE
end

function CSheet:GetColDefaults()
    return self.sheet.DEFAULT
end

function CSheet:GetData()
    return self.sheet.DATA
end

function CSheet:ReadCell(row, col)
    return self.sheet.DATA[row][col]
end

function CSheet:GetColCnt()
    return self.sheet.COL_CNT
end

function CSheet:GetRowCnt()
    return self.sheet.ROW_CNT
end