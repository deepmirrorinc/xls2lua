function checkINT(data)
    if type(data) ~= "number" then
        data = tonumber(data)
        if data == nil then
            data = 0
        end
    end
    return true, math.floor(data)
end

return checkINT