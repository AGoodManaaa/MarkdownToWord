-- 在输出前把所有数学节点改成普通文本，保留 $$ / $ 包裹，
-- 同时把公式内部的换行/缩进压缩成一个空格，方便后处理。

-- 简单的空白规范化：去掉首尾空白，内部空白统一为一个空格
local function normalize_tex(s)
  -- 去掉开头/结尾的所有空白（包括换行、空格、tab）
  s = s:gsub("^%s+", ""):gsub("%s+$", "")
  -- 把中间任意长度的空白（空格/换行/tab）压成一个空格
  s = s:gsub("%s+", " ")
  return s
end

local function math_to_text(el)
  local open_delim, close_delim

  if el.mathtype == "DisplayMath" then
    -- 块公式：用 $$ 包裹
    open_delim, close_delim = "$$", "$$"
  else
    -- 行内公式：用 $ 包裹
    open_delim, close_delim = "$", "$"
  end

  -- el.text 是不含分隔符的 LaTeX 内容
  local content = normalize_tex(el.text)
  local s = open_delim .. content .. close_delim

  return pandoc.Str(s)
end

return {
  {
    Math = math_to_text
  }
}
