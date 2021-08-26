-- File: cfgreader.lua
-- Desc:

CCfgReader = Class("CCfgReader")

function CCfgReader:CCfgReader(sheet)
    self.sheet = sheet
    self.row = 1
    self.maxCol = sheet:GetColCnt()
    self.colIdx = 0
    self.types = sheet:GetColTypes()
    self.defaults = sheet:GetColDefaults()
    self.fields = sheet:GetLogicNames()
end

function CCfgReader:SetRow(row)
    self.row = row
    self.colIdx = 0
end

function CCfgReader:ReadNext()
    self.colIdx = self.colIdx + 1
    if self.colIdx > self.maxCol then
        return nil
    end
    local val = self.sheet:ReadCell(self.row, self.colIdx)
    if nil == val then
        return self.defaults[self.types[self.colIdx]]
    end
    return val
end

function CCfgReader:GetFields()
    return self.fields
end

function CCfgReader:Readline(itemInst)
    for i, name in ipairs(self.fields) do
        itemInst[name] = self:ReadNext()
    end
end

function CCfgReader:ReadInt()
    return self:ReadNext()
end

function CCfgReader:ReadFloat()
    return self:ReadNext()
end

function CCfgReader:ReadString()
    return self:ReadNext()
end

