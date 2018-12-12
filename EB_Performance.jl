
include("../../credentials.jl")
include("dirs.jl")


using Cache
using SQLite
using EBDB
using XlsxWriter
using ExcelReaders
using DataArrays

maps = Dict("EB"=>Dict(

dir = raw"N:\EB Performance Sheets"

function perfXls(c)
	for xlfn in sort(filter((f)->f[1] != '~' && f[end-4:end] == ".xlsx" && f[1:4]=="2018", readdir(dir)))
		file_recorded(xlfn) || put!(c, (dir, xlfn))
	end
end



function store_Paint_sheet(inum, xld, xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "Paint")
	vals = Vector{Any}(25)
	ws_n = Dict("Shift"=>1, "Product"=>2, "Std_Rate_PPH"=>3, "Avail_Hours"=>4, "Product_Max"=>5, "Product_Avail"=>6, "Product_Variance"=>7, "Time_Variance"=>8, "Efficiency"=>9, "Line"=>10, "Item"=>11, "Process"=>12, "Problem"=>13, "Quality_Defect_Type"=>14, "Quality_Lost"=>15, "Loss_Mins"=>16, "Effect"=>17, "OEE_Element"=>18, "Action"=>19, "Fix_Or_Repair"=>20)
	
	val_n = Dict("Shift"=>3, "Product"=>4, "Std_Rate_PPH"=>5, "Avail_Hours"=>6, "Product_Max"=>7, "Product_Avail"=>8, "Product_Variance"=>9, "Time_Variance"=>10, "Efficiency"=>11, "Line"=>12, "Item"=>13, "Process"=>14, "Problem"=>15, "Quality_Defect_Type"=>16, "Quality_Lost"=>17, "Loss_Mins"=>18, "Effect"=>19, "OEE_Element"=>20, "Action"=>21, "Fix_Or_Repair"=>22)
	

	Data(r, n::Integer) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	Data(r, s::String) = Data(r, ws_n[s])
	vv(s) = vals[val_n[s]]
	dv(r, s) = data[r, ws_n[s]]
	vdif!(r, s) = vals[val_n[s]] = Data(r, s) ? dv(r, s) : vv(s)
	vd!(r, s) = vals[val_n[s]] = dv(r, s)
	vfloor!(r, s) = 
	
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,19]
	
	for r in 3:50
		if Data(r, 11)
			println(data[r,1:end])
		end
	
		if !Data(r, 11)
			return inum
		end
		vdif!(r, "Shift")
		vdif!(r, "Product")
		
		if Data(r, ws_n["Std_Rate_PPH"])
			vd!("Std_Rate_PPH")
			for s in ["Avail_Hours", "Product_Max", "Product_Avail", "Product_Variance", "Time_Variance", "Efficiency"]
				vdif!(r, s)
			end
			vals[vals_n["Time_Variance"]] = floor(data[r, ws_n["Time_Variance"]])
			vals[vals_n["Efficiency"]] = floor(data[r, ws_n["Efficiency"]])
		else
			if Data(r, ws_n["Product"])
				for k = 6:12
					vals[k] = 0
				end
			end
		end
		for k in 11:19
			if Data(r, k) # Item
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
		vals[23] = inum
		vals[24] = xlfn
		insertPaint(vals)
	end
	inum
end
	


function store_EB_sheet(inum, xld, xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "EB Line")
	vals = Vector{Any}(23)
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,18]
	
	Data(r, n) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	
	for r in 3:50
		Data(r, 11) && (println(data[r,1:end])) || (return inum)
		
		for c in 1:3
			Data(r,c) && vals[c+2] = data[r, c] # shift
		end
		if Data(r,4) # pph
			vals[6] = data[r, 4]
			for c in 5:12				
				Data(r,c) && (vals[c+2] = data[r, c])
			end
			vals[8] = floor(data[r, 6])
			vals[10] = floor(data[r, 8])
		else
			Data(r,3) && (vals[6:12] = 0)
		end
		for c in 11:19
			if Data(r,c)
				vals[c+2] = data[r, c]
			else
				vals[c+2] = k == 14 ? 0:""
			end
		end

		inum += 1
		vals[22] = inum
		vals[23] = xlfn
		insertEB(vals)
	end
	inum
end

function store_sheets()
	inum = last_inum()
	for (xld, xlfn) in Channel(perfXls)
		inum = store_sheet(inum, xld, xlfn)
	end
end


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

#clear("EB")
store_sheets()
open(write_sheets, raw"N:\EB Performance Sheets\consolidated.txt", "w+")
