
Import "xls2lua.sheet"

CXls = Class("CXls")

function CXls:CXls()
    self.xls = luacom.GetObject("Excel.Application") or luacom.CreateObject("Excel.Application")
    assert(self.xls)
    self.xls.Visible = false
    self.xls.Application.DisplayAlerts = false
    self.book = nil
    self.sheets = nil                -- {name = true/sheetObj}
end

function CXls:Open(path)
    if not self.FileExists(path) then
        print(string.format("File NOT exists! Path:", path))
        return false
    end

    self.book = self.xls.WorkBooks:Open(path, nil, ReadOnly)

    local enum = luacom.GetEnumerator(self.xls.Sheets)
    local sh = enum:Next()
    self.sheets = {}
    while sh do
        local sheet = sheet.CSheet(self, sh)
        if sheet:IsValid() then
            self.sheets[sh.Name]  = sheet
        end
        sh = enum:Next()
    end
    return true
end

function CXls:CloseBook()
    if self.book then
        for k, sheet in pairs(self.sheets) do
            self.sheets[k] = nil
        end
        self.sheets = nil
        self.book:Close()
        self.book = nil
    end
end

function CXls:Close()
    self:CloseBook()
    if self.xls then
        self.xls.Application:Quit()
    end
end

function CXls:GetSheets()
    return self.sheets
end

function CXls:ReadCell(row, col)
    return self.xls.ActiveSheet.Cells(row, col).Value2
end

function CXls.FileExists(path)
    local fp = io.open(path, "r")
    if fp then
        fp:close()
        return true
    else
        return false
    end
end