
include("../../credentials.jl")
include("dirs.jl")


using Cache
using SQLite
using EBDB
using XlsxWriter
using ExcelReaders
using DataArrays

exit()

dir = raw"N:\EB Performance Sheets"

function perfXls(c)
	for xlfn in sort(filter((f)->f[1] != '~' && f[end-4:end] == ".xlsx" && f[1:4]=="2018", readdir(dir)))
		Cache.validate_by_mtime!("$dir\\$xlfn", "EBP_$(xlfn)") && put!(c, "$dir\\$xlfn")
	end
end

function store_sheet(inum, ins, data)
	SQLite.bind!(ins, 1, Dates.value(data[1,2]))
	SQLite.bind!(ins, 2, data[1,18])
	
	for r in 3:50
		if typeof(data[r,11]) != DataValues.DataValue{Union{}}
			println(data[r,1:end])
		end
	
		if typeof(data[r,11]) == DataValues.DataValue{Union{}}
			return inum
		end
		for k in 1:3
			if typeof(data[r,k]) != DataValues.DataValue{Union{}} # shift
				SQLite.bind!(ins, k+2, data[r, k])
			end
		end
		if typeof(data[r,4]) != DataValues.DataValue{Union{}} # pph
			SQLite.bind!(ins, 6, data[r, 4])
			for k in 5:12				
				if typeof(data[r,k]) != DataValues.DataValue{Union{}} # avail hrs
					SQLite.bind!(ins, k+2, data[r, k])
				end
			end
			SQLite.bind!(ins, 8, floor(data[r, 6]))
			SQLite.bind!(ins, 10, floor(data[r, 8]))
		else
			if typeof(data[r,3]) != DataValues.DataValue{Union{}}
				for k = 6:12
					SQLite.bind!(ins, k, 0)
				end
			end
		end
		for k in 11:19
			if typeof(data[r,k]) != DataValues.DataValue{Union{}} # Item
				SQLite.bind!(ins, k+2, data[r, k])
			else
				if k == 14
					SQLite.bind!(ins, k+2, 0)
				else
					SQLite.bind!(ins, k+2, "")
				end
			end
		end

		inum += 1
		SQLite.bind!(ins, 22, inum)
		SQLite.execute!(ins)
	end
	inum
end

function store_sheets()
	SQLite.query(db, "delete from EB")
	ins = SQLite.Stmt(db, "Insert into EB (Date, Leader, Shift, Line, Product, Std_Rate_PPH, Avail_Hours, Product_Max, Product_Avail, Product_Variance, Time_Variance, Efficiency, Item, Process, Problem, Loss_Mins, Effect, OEE_Element, Action, Fix_Or_Repair, Weld_section_due, IncidentNo) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

	inum = 0
	for xlfn in Channel(perfXls)
		println(xlfn)
		inum = store_sheet(inum, ins, readxlsheet(xlfn, "EB Line"))		
	end
end

store_sheets()

function write_sheets(io)
	rows = SQLite.query(db, "select * from EB")
	cells = size(rows[1,1:end],2)-1
	for row in 1:size(rows,1)
		print(io, rows[row, 22])
		print(io, "\t", Dates.format(DateTime(Dates.UTM(rows[row,1])), "yyyy-mm-dd"))
		for c in 2:cells
			print(io, "\t", rows[row,c])
		end
		println(io)
	end
end

open(write_sheets, raw"N:\EB Performance Sheets\consolidated.txt", "w+")
