--
-- 配置说明：
--  RAW_PATH表示原始数据的路径，结构如下
--              |---datacls（文件夹名字必须一致）导表过程中的一些自定义数据检测，自动生成
--  RAW_PATH---|                                |---xls2lua_config.lua文件，包含所有的xlsx文件名字
--              |---excel（文件夹名字必须一致）---|
--                                             |---导表的原始xlsx文件
--
--  LUA_PATH表示新生成的lua table数据

RAW_PATH = "D:/PokemonBattle/xls2lua/Assets/Resources/data/rawdata"
LUA_PATH = "D:/PokemonBattle/xls2lua/Assets/Resources/data/luadata"