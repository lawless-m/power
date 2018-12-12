
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
		file_recorded(xlfn) || put!(c, (dir, xlfn))
	end
end



function store_Paint_sheet(inum, xld, xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "Paint")
	vals = Vector{Any}(25)
	ws_n = Dict("Shift"=>1, "Product"=>2, "Std_Rate_PPH"=>3, "Avail_Hours"=>4, "Product_Max"=>5, "Product_Actual"=>6, "Product_Variance"=>7, "Time_Variance"=>8, "Efficiency"=>9, "Line"=>10, "Item"=>11, "Process"=>12, "Problem"=>13, "Quality_Defect_Type"=>14, "Quality_Lost_Parts"=>15, "Loss_Mins"=>16, "Effect"=>17, "OEE_Element"=>18, "Action"=>19, "Fix_Or_Repair"=>20)
	
	val_n = Dict("Shift"=>3, "Product"=>4, "Std_Rate_PPH"=>5, "Avail_Hours"=>6, "Product_Max"=>7, "Product_Actual"=>8, "Product_Variance"=>9, "Time_Variance"=>10, "Efficiency"=>11, "Line"=>12, "Item"=>13, "Process"=>14, "Problem"=>15, "Quality_Defect_Type"=>16, "Quality_Lost_Parts"=>17, "Loss_Mins"=>18, "Effect"=>19, "OEE_Element"=>20, "Action"=>21, "Fix_Or_Repair"=>22)
	

	Data(r, n::Integer) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	Data(r, s::String) = Data(r, ws_n[s])
	vv(s) = vals[val_n[s]]
	dv(r, s) = data[r, ws_n[s]]
	vdif!(r, s) = vals[val_n[s]] = Data(r, s) ? dv(r, s) : vv(s)
	vd!(r, s) = vals[val_n[s]] = dv(r, s)
	vfloor!(r, s) = vals[vals_n[s]] = floor(dv(r, s))
	
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,19]
	
	for r in 3:50
		if !Data(r, ws_n["Item"])
			return inum
		end
		
		println(data[r,1:end])
		
		vdif!(r, "Shift")
		vdif!(r, "Product")
		
		if Data(r, ws_n["Std_Rate_PPH"])
			vd!("Std_Rate_PPH")
			for s in ["Avail_Hours", "Product_Max", "Product_Actual", "Product_Variance"]
				vdif!(r, s)
			end
			vfloor!(r, "Time_Variance")
			vfloor!(r, "Efficiency")
		else
			if Data(r, ws_n["Product"])
				for k = 6:12
					vals[k] = 0
				end
			end
		end
		for k in 11:19
			if Data(r, ws_n) # Item
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
	
	
	ws_n = Dict("Shift"=>1, "Line"=>2, "Product"=>3, "Std_Rate_PPH"=>4, "Avail_Hours"=>5, "Product_Max"=>6, "Product_Actual"=>7, "Product_Variance"=>8, "Time_Variance"=>9, "Efficiency"=>10, "Item"=>11, "Process"=>12, "Problem"=>13, "Loss_Mins"=>14, "Effect"=>15, "OEE_Element"=>16, "Action"=>17, "Fix_Or_Repair"=>18, "Weld_Section_Due"=>19)
	
	val_n = Dict("Shift"=>3, "Line"=>4, "Product"=>5, "Std_Rate_PPH"=>6, "Avail_Hours"=>7, "Product_Max"=>8, "Product_Actual"=>9, "Product_Variance"=>10, "Time_Variance"=>11, "Efficiency"=>12, "Item"=>13, "Process"=>14, "Problem"=>15, "Loss_Mins"=>16, "Effect"=>17, "OEE_Element"=>18, "Action"=>19, "Fix_Or_Repair"=>20, "Weld_Section_Due"=>21, "IncidentNo"=>22, "Filename"=>23)
	
	

	Data(r, n::Integer) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	Data(r, s::String) = Data(r, ws_n[s])
	vv(s) = vals[val_n[s]]
	dv(r, s) = data[r, ws_n[s]]
	vdif!(r, s) = vals[val_n[s]] = Data(r, s) ? dv(r, s) : vv(s)
	vd!(r, s) = vals[val_n[s]] = dv(r, s)
	vfloor!(r, s) = vals[vals_n[s]] = floor(dv(r, s))
	v!(s, v) = vals[val_n[s]] = v
	
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,18]
	
	Data(r, n) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	
	for r in 3:50
		if !Data(r, ws_n["Item"])
			return inum
		end
		
		println(data[r,1:end])
		
		vdif!(r, "Shift")
		vdif!(r, "Line")
		vdif!(r, "Product")
		
		if Data(r,ws_n["Std_Rate_PPH"]) # pph
			vd!(r, "Std_Rate_PPH")
			foreach((s)->vdif!(r, s), ["Avail_Hours", "Product_Max", "Product_Actual", "Time_Variance", "Item", "Process"])
			vfloor!(r, "Product_Variance")
			vfloor!(r, "Efficiency")
		else
			if Data(r,ws_n["Product"])
				vals[[ws_n[s] for s in ["Product_Max", "Product_Actual", "Product_Variance", "Time_Variance", "Efficiency", "Item", "Process"]]] = 0
			end
		end
		for s in ["Item", "Process", "Problem", "Loss_Mins", "Effect", "OEE_Element", "Action", "Fix_Or_Repair", "Weld_Section_Due"]
			if Data(r,s)
				vd!(r, s)
			else
				v!(s, "Loss_Mins" ? 0 : "")
			end
		end

		inum += 1
		v!("IncidentNo", inum)
		v!("Filename", xlfn)
		insertEB(vals)
	end
	inum
end

function store_EB_sheets()
	inum = last_inum()
	for (xld, xlfn) in Channel(perfXls)
		inum = store_EB_sheet(inum, xld, xlfn)
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
store_EB_sheets()
open(write_sheets, raw"N:\EB Performance Sheets\consolidated.txt", "w+")
