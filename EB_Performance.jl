
include("../../credentials.jl")
include("dirs.jl")


using Cache
using SQLite
using EBDB
using XlsxWriter
using ExcelReaders
using DataArrays


dir = raw"N:\EB Performance Sheets"

function perfXls(c)
	for xlfn in sort(filter((f)->f[1] != '~' && f[end-4:end] == ".xlsx" && f[1:4]=="2018", readdir(dir)))
		put!(c, (dir, xlfn))
	end
end

function store_sheet(inum, xlfn)
	data = readxlsheet(xlfn, "EB Line")
	vals = Vector{Any}(22)
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,18]
	
	for r in 3:50
		if typeof(data[r,11]) != DataValues.DataValue{Union{}}
			println(data[r,1:end])
		end
	
		if typeof(data[r,11]) == DataValues.DataValue{Union{}}
			return inum
		end
		for k in 1:3
			if typeof(data[r,k]) != DataValues.DataValue{Union{}} # shift
				vals[k+2] = data[r, k]
			end
		end
		if typeof(data[r,4]) != DataValues.DataValue{Union{}} # pph
			vals[6] = data[r, 4]
			for k in 5:12				
				if typeof(data[r,k]) != DataValues.DataValue{Union{}} # avail hrs
					vals[k+2] = data[r, k]
				end
			end
			vals[8] = floor(data[r, 6])
			vals[10] = floor(data[r, 8])
		else
			if typeof(data[r,3]) != DataValues.DataValue{Union{}}
				for k = 6:12
					vals[k] = 0
				end
			end
		end
		for k in 11:19
			if typeof(data[r,k]) != DataValues.DataValue{Union{}} # Item
				vals[k+2] = data[r, k]
			else
				if k == 14
					vals[k+2] = 0
				else
					vals[k+2] = ""
				end
			end
		end

		inum += 1
		vals[22] = inum
		insertEB(vals)
	end
	inum
end

function store_sheets()
	inum = 0
	for (xld, xlfn) in Channel(perfXls)
		inum = cache(xlfn, ()->store_sheet(inum, "$xld\\$xlfn"), "EBPerf")
	end
end

store_sheets()

function write_sheets(io)
	rows = allEB()
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
