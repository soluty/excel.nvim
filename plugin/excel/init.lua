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

function RenderTableTest()
	local NuiTable = require("nui.table")
	local Text = require("nui.text")
	local tbl = NuiTable({
		bufnr = 0,
		columns = {
			{
				align = "center",
				header = "Name",
				columns = {
					{ accessor_key = "firstName", header = "First" },
					{
						id = "lastName",
						accessor_fn = function(row)
							return row.lastName
						end,
						header = "Last",
					},
				},
			},
			{
				align = "right",
				accessor_key = "age",
				cell = function(cell)
					return Text(tostring(cell.get_value()), "DiagnosticInfo")
				end,
				header = "Age",
			},
		},
		data = {
			{ firstName = "John", lastName = "Doe", age = 42 },
			{ firstName = "Jane", lastName = "Doe", age = 27 },
		},
	})

	tbl:render()
end
