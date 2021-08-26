function checkIntRange(data)
    local parts = {}
    if type(data) == "table" then
        parts = data
    else
        for v in data:gmatch("[^,]+") do
            table.insert(parts, v)
        end
    end
    
    if #parts ~= 2 then
        print ("parts number", #parts, parts[1], parts[2])
        return false, "IntRange类型应该有一个逗号将两个整数分隔"
    end
    
    for i, part in pairs(parts) do
        local v 
        if type(part) == "number" then
            v = part
        else
            local s = (parts[i]:gsub("^%s*(.-)%s*$", "%1"))
            v = tonumber(s)
        end
        
        if v == nil or math.floor(v) ~= v then
            print("value =", v, s)
            return false, string.format("第%d个数值必须为整数", i)
        end
        parts[i] = v
    end
    
    if parts[1] > parts[2] then
        return false, "第一个数必须小于等于第二个数"
    end
    
    return true, parts
end

return checkIntRange
