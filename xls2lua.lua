
local Sheet = require 'Sheet'
dofile("../../../")

-------------------------------------------------------------
-- class Excel
local Excel = {}

function Excel.new()
	local o = {}
	setmetatable(o, Excel)
	Excel.__index = Excel
	return o
end

function Excel:open(path, visible)
	if not self:isExist(path) then
		return false
	end

	local excel = luacom.GetObject('Excel.Application') or luacom.CreateObject('Excel.Application')
	--是否显示excel窗口
	excel.Visible = visible
	assert(excel)
	self.excel = excel
	
	excel.Application.DisplayAlerts = false
	local book = excel.WorkBooks:Open(G2U(path),nil,ReadOnly)
	--book.Saved = false
	self.book = book
	
	return true
end

function Excel:save( )
	if self.book then
		self.book:Save()
	end
end

function Excel:close()
	if self.book then
		self.book:Close()
	end
	if self.excel then
		self.excel.Application:Quit()
	end
end

function Excel:getSheet(name)
	local e = luacom.GetEnumerator(self.excel.Sheets)
	local sheet = e:Next()
	while sheet do
		if sheet.Name == name then
			return Sheet.new(sheet, self)
		end
		sheet = e:Next()
	end
	return nil
end

function Excel:createSheet(name)
	local sheet = self:getSheet(name)
	--没有则创建
	if sheet then
		return sheet
	end
	sheet = self.book.Sheets:Add()
	sheet.name = name
	return Sheet.new(sheet, self)
end

function Excel:isExist(path)
	local t=io.open(path, 'r')
	if not t then
		return false
	else
	 	t:close()
		return true
	end
end

return Excel
