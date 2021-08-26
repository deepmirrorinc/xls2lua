function checkFLOAT(data)
    if type(data) ~= "number" then
        data = tonumber(data)
        if data == nil then
            data = 0.0
        end
    end
    return true, data
end

return checkFLOAT
