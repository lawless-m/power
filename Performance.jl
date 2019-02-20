2
include("../../credentials.jl")
include("dirs.jl")

using Cache
using SQLite
using PerfDB
using XlsxWriter
using ExcelReaders
using DataArrays

WS_Cols = Dict("EB"=>Dict("Shift"=>1, "Line"=>2, "Product"=>3, "Std_Rate_PPH"=>4, "Avail_Hours"=>5, "Product_Max"=>6, "Product_Actual"=>7, "Product_Variance"=>8, "Time_Variance"=>9, "Efficiency"=>10, "Item"=>11, "Process"=>12, "Problem"=>13, "Loss_Mins"=>14, "Effect"=>15, "OEE_Element"=>16, "Action"=>17, "Fix_Or_Repair"=>18, "Weld_Section_Due"=>19),
"Paint"=>Dict("Shift"=>1, "Product"=>2, "Std_Rate_PPH"=>3, "Avail_Hours"=>4, "Product_Max"=>5, "Product_Actual"=>6, "Product_Variance"=>7, "Time_Variance"=>8, "Efficiency"=>9, "Line"=>10, "Item"=>11, "Process"=>12, "Problem"=>13, "Quality_Defect_Type"=>14, "Quality_Lost_Parts"=>15, "Loss_Mins"=>16, "Effect"=>17, "OEE_Element"=>18, "Action"=>19, "Fix_Or_Repair"=>20))

DB_Cols = Dict("EB"=>Dict("Date"=>1, "Leader"=>2, "Shift"=>3, "Line"=>4, "Product"=>5, "Std_Rate_PPH"=>6, "Avail_Hours"=>7, "Product_Max"=>8, "Product_Actual"=>9, "Product_Variance"=>10, "Time_Variance"=>11, "Efficiency"=>12, "Item"=>13, "Process"=>14, "Problem"=>15, "Loss_Mins"=>16, "Effect"=>17, "OEE_Element"=>18, "Action"=>19, "Fix_Or_Repair"=>20, "Weld_Section_Due"=>21, "IncidentNo"=>22, "Filename"=>23),
"Paint"=>Dict("Date"=>1, "Leader"=>2, "Shift"=>3, "Product"=>4, "Std_Rate_PPH"=>5, "Avail_Hours"=>6, "Product_Max"=>7, "Product_Actual"=>8, "Product_Variance"=>9, "Time_Variance"=>10, "Efficiency"=>11, "Line"=>12, "Item"=>13, "Process"=>14, "Problem"=>15, "Quality_Defect_Type"=>16, "Quality_Lost_Parts"=>17, "Loss_Mins"=>18, "Effect"=>19, "OEE_Element"=>20, "Action"=>21, "Fix_Or_Repair"=>22, "IncidentNo"=>23, "Filename"=>24))


function store_Paint_sheet(inum, xld, xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "Paint")
	println(xlfn)

	ws_n = WS_Cols["Paint"]
	val_n = DB_Cols["Paint"]

	vals = Vector{Any}(length(val_n))

	Data(r, n::Integer) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	Data(r, s::String) = Data(r, ws_n[s])
	vv(s) = vals[val_n[s]]
	dv(r, s) = data[r, ws_n[s]]
	vdif!(r, s) = v!(s, Data(r, s) ? dv(r, s) : vv(s))
	vd!(r, s) = v!(s, dv(r, s))
	vround!(r, s) = v!(s, convert(Int, round(dv(r, s), RoundNearestTiesUp)))
	vroundz!(r, s) = v!(s, convert(Int, round(vfifnot!(r, s), RoundNearestTiesUp)))
	v!(s, v) = vals[val_n[s]] = v
	vi!(s, r) = vi!(s, convert(Int, dv(r, s)))
	vz!(s::String) = vals[val_n[s]] = 0
	vz!(vs) = foreach(s->vals[val_n[s]] = 0, vs)

	vfifnot!(r, s) = v!(s, Data(r, s) ? dv(r, s) : 0.0)

	vzifnot!(r, s) = v!(s, Data(r, s) ? (dv(r, s) isa Number ? dv(r, s) : 0) : 0)
	vesifnot!(r, s) = v!(s, Data(r, s) ? dv(r, s) : "")

	println(data[1,1:end])
	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,19]

	for r in 3:50
		if (!Data(r, ws_n["Item"])) && (!Data(r, ws_n["Quality_Defect_Type"]))
			return inum
		end

		println(data[r,1:end])

		vdif!(r, "Shift")
		vdif!(r, "Product")

		if Data(r, ws_n["Std_Rate_PPH"])
			foreach((s)->vround!(r, s), ["Std_Rate_PPH", "Product_Max", "Product_Variance"])
			vroundz!(r, "Product_Actual")
			foreach((s)->vdif!(r, s), ["Avail_Hours", "Time_Variance", "Efficiency"])
		else
			if Data(r,ws_n["Product"])
				vz!(["Std_Rate_PPH", "Product_Max", "Product_Actual", "Product_Variance", "Avail_Hours", "Time_Variance", "Efficiency"])
			end
		end
		for s in ["Line", "Item", "Process", "Problem", "Loss_Mins", "Effect", "OEE_Element", "Action", "Fix_Or_Repair", "Quality_Defect_Type",  "Quality_Lost_Parts"]
			vesifnot!(r, s)
		end
		for s in ["Loss_Mins",  "Quality_Lost_Parts"]
			vzifnot!(r, s)
		end

		inum += 1
		vals[23] = inum
		vals[24] = xlfn

		#srt(a, b) = val_n[a] < val_n[b]

		#for s in sort(collect(keys(val_n)), lt=srt)
		#	println(val_n[s], " - ", s, " - ", vals[val_n[s]])
		#end

		insertPaint(vals)
	end
	inum
end

function store_EB_sheet(inum, xld, xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "EB Line")
	println(xlfn)

	ws_n = WS_Cols["EB"]
	val_n = DB_Cols["EB"]

	vals = Vector{Any}(length(val_n))

	Data(r, n::Integer) = typeof(data[r, n]) != DataValues.DataValue{Union{}}
	Data(r, s::String) = Data(r, ws_n[s])
	vv(s) = vals[val_n[s]]
	dv(r, s) = data[r, ws_n[s]]
	vdif!(r, s) = v!(s, Data(r, s) ? dv(r, s) : vv(s))
	vd!(r, s) = v!(s, dv(r, s))
	vround!(r, s) = v!(s, convert(Int, round(dv(r, s), RoundNearestTiesUp)))
	v!(s, v) = vals[val_n[s]] = v
	vi!(s, r) = vi!(s, convert(Int, dv(r, s)))
	vz!(s::String) = vals[val_n[s]] = 0
	vz!(vs) = foreach(s->vals[val_n[s]] = 0, vs)
		vzifnot!(r, s) = v!(s, Data(r, s) ? (dv(r, s) isa Number ? dv(r, s) : 0) : 0)
	vesifnot!(r, s) = v!(s, Data(r, s) ? dv(r, s) : 0)

	vals[1] = Dates.value(data[1,2])
	vals[2] = data[1,18]

	for r in 3:50
		if !Data(r, ws_n["Item"])
			return inum
		end

		println(data[r,1:end])

		vdif!(r, "Shift")
		vdif!(r, "Line")
		vdif!(r, "Product")

		if Data(r,ws_n["Std_Rate_PPH"]) # pph
			foreach((s)->vround!(r, s), ["Std_Rate_PPH", "Product_Max", "Product_Actual", "Product_Variance"])
			foreach((s)->vdif!(r, s), ["Avail_Hours", "Time_Variance", "Efficiency"])
		else
			Data(r,ws_n["Product"]) && vz!(["Std_Rate_PPH", "Product_Max", "Product_Actual", "Product_Variance", "Avail_Hours", "Time_Variance", "Efficiency"])
		end
		foreach(s-> Data(r,s) ? vd!(r, s) : v!(s, s=="Loss_Mins" ? 0 : ""), ["Item", "Process", "Problem", "Loss_Mins", "Effect", "OEE_Element", "Action", "Fix_Or_Repair", "Weld_Section_Due"])

		inum += 1
		v!("IncidentNo", inum)
		v!("Filename", xlfn)
		insertEB(vals)
	end
	inum
end

function write_sheets(io, dfn, line)
	rows = dfn()
	cells = DB_Cols[line]["IncidentNo"]
	for row in 1:size(rows,1)
		print(io, rows[row, 1])
		print(io, "\t", Dates.format(DateTime(Dates.UTM(rows[row,2])), "yyyy-mm-dd"))
		foreach(c->print(io, "\t", rows[row,c]), 3:cells)
		println(io)
	end
end

function perfXls(c, table, shtdir)
	for xlfn in sort(filter((f)->f[1] != '~' && f[end-4:end] == ".xlsx" && f[1:3]=="201", readdir(shtdir)))
		file_recorded(xlfn, table) || put!(c, (shtdir, xlfn))
	end
end

function store_sheets(line, xlfun, storefn)
	inum = last_inum(line)
	for (xld, xlfn) in Channel(xlfun)
		inum = storefn(inum, xld, xlfn)
	end
end

function procEB()
	dir = "N:\\EB Performance Sheets"
	#clear("EB")
	store_sheets("EB", (c)->perfXls(c, "EB", dir), store_EB_sheet)
	open((io)->write_sheets(io, allEB, "EB"), "$dir\\EB consolidated.txt", "w+")
end

function procPnt()
	dir = "N:\\Paint Performance Sheets"
	#clear("Paint")
	store_sheets("Paint", (c)->perfXls(c, "Paint", dir), store_Paint_sheet)
	open((io)->write_sheets(io, allPaint, "Paint"), "$dir\\Paint consolidated.txt", "w+")
end

procPnt()
#procEB()
