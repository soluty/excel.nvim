vim.cmd([[
  if exists('g:loaded_excelnvim')
    finish
  endif
  let g:loaded_excelnvim = 1

  function! s:RequireExcel(host) abort
    return jobstart(['excelnvim'], {'rpc': v:true})
  endfunction

  call remote#host#Register('excelnvim', 'x', function('s:RequireExcel'))

  call remote#host#RegisterPlugin('excelnvim', '0', [
      \ {'type': 'function', 'name': 'ExcelNvimOpenFile', 'sync': 1, 'opts': {}},
      \ ])
]])

local M = {}

local function convertToBase26(i)
	local result = ""
	while i > 0 do
		local remainder = (i - 1) % 26
		result = string.char(remainder + 65) .. result
		i = math.floor((i - 1) / 26)
	end
	return result
end

-- M.ExcelNvimShowTable = function ()
function ExcelNvimShowTable()
	local json = require("json")
	local fileName = vim.fn.expand("%:p")
	local resultStr = vim.api.nvim_exec(string.format("let ret = ExcelNvimOpenFile('%s') | echo ret", fileName), true)
	local NuiTable = require("nui.table")
	local Text = require("nui.text")
	resultStr = string.gsub(resultStr, "'", '"')
	local result = json.decode(resultStr)
	local tblData = result.Tables[1]
	local maxCol = tblData.MaxCol
	local columns = {}
	for i = 1, maxCol do
		table.insert(columns, {
			header = convertToBase26(i),
			accessor_key = i,
		})
	end
	vim.cmd("tabedit")
	local tbl = NuiTable({
		bufnr = 0,
		columns = columns,
		data = tblData.Cells,
	})
	tbl:render()
end

return M
