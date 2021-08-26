function checkTABLE(data)
    if data == "" or data == nil then
        return true, {}
    elseif type(data) == "table" then
        return true, data
    else
        local code = "D="..data.."\nreturn D" 
        local func = loadstring(code)
        local bOK, ret = pcall(func)
        if not bOK then
            return false, string.format("INVIAL lua table, may brackets MISMATCH! cell data=\"%s\"", data)
        end
        return true, ret
    end
end

return checkTABLE
