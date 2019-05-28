
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
	println(xlfn)
	data = readxlsheet(xld * "\\" * xlfn, "EB Line")

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

function eb_causes(a, b)
	cause = ""
	if a in ["Bar code", "barcode labeller", "barcode ", "Barcode rejects", "Label Printer", "Label printer"]
		cause = "Barcode printer"
	elseif a in ["Briefing", "briefing"]
		cause = "Planned"
	elseif a in ["Bay clean", "Bay Clean"]
		cause = "PPM"
	elseif a == "Build"
		if b == "parts"
			cause = "No Parts"
		elseif b == "SECTION"
			cause = "Section"
		elseif b == "set up"
			cause = "Changeover"
		elseif b == "tooling"
			cause = "Tooling"
		else
			cause = "Build"
		end
	elseif a in ["Camera", "camera"]
		cause = "Camera"
	elseif a in ["Cell Exit", "Changeover"]
		if b == "Run Out"
			cause = "Run Out"
		elseif b == "Filament"
			cause = "Filament"
		elseif b in ["section", "SECTION", "Section", "Secton"]
			cause = "Section"
		else
			cause = "Changeover"
		end
	elseif a == "Cell Load"
		if b == "cold start"
			cause = "Cold Start"
		elseif b == "line fill"
			cause = "Line Fill"
		elseif b == "process"
			cause = "Line Fill"
		else
			cause = "Cold Start"
		end
	elseif a == "Cell runout"
		cause = "Run out"
	elseif a == "Changeover"
		if b == "cell Runout"
			cause = "Run out"
		elseif b == "Filament"
			cause = "Filament"
		elseif b == "line fill"
			cause = "Line fill"
		elseif b == "Run Out"
			b = "Run Out"
		elseif b in ["section", "SECTION", "Section"]
			cause = "Section"
		else
			cause = "Changeover"
		end
	elseif a == "Checks"
		cause = "TPM"
	elseif a == "Daily pms"
		if b == "cold start"
			cause = "Cold Start"
		else
			cause = "TPMs"
		end
	elseif a in ["Cold start", "Cold Start"]
		cause = "Cold Start"
	elseif a == "Domino printer"
		cause = "Barcode printer"
	elseif a == "Dot Matrix"
		cause = "Dot Matrix"
	elseif a == "Elbow"
		cause = "Planned"
	elseif a == "Engineer on line"
		cause = "Planned"
	elseif a == "Filament"
		if b in ["SECTION", "Secton", "Section"]
			cause = "Section"
		else
			cause = "Filament"
		end
	elseif a == "Fill"
		if b == "tooling"
			cause = "Tooling"
		elseif b == "vacc"
			cause = "Filling vacc"
		end
	elseif a == "Finish & RP"
		cause = "No Parts"
	elseif a == "Cell Runout"
		cause = "Runout"
	elseif a == "Filament"
		if b in ["SECTION", "Section", "Secton"]
			cause = "Section"
		else
			cause = "Filament"
		end
	elseif a == "Fill"
		if b in ["Silicone Change","Silicone change"]
			cause = "Silicone Change"
		end
	elseif a == "Hare Press"
		cause = "Press"
	elseif a == "Inertia"
		cause = "Inertia"
	elseif a == "Leak & Volume"
		cause = a
	elseif a == "Line Fill"
		cause = a
	elseif a in ["Load bearing", "Load Bearing"]
		cause = "Tooling"
	elseif a in ["Man", "manning", "manpower"]
		if b == "Silicone change"
			cause = b
		else
			cause = "Operator shortage"
		end
	elseif a == "meeting"
		cause = "Planned"
	elseif a in ["P Stamp", "P-Stamp"]
		cause = "Tooling"
	elseif a == "Pack"
		cause = "Packing"
	elseif a in ["PPM", "ppms", "PPM's"]
		cause = "PPM"
	elseif a in ["printer", "Printer"]
		cause = "Barcode printer"
	elseif a == "Rework Loop"
		if b == "Run Out"
			cause = "Run Out"
		else
			cause = "Reworks"
		end
	elseif a == "Robot 1"
		cause = "Robot 1"
	elseif a == "Robot 2"
		cause = "Robot 2"
	elseif a == "Roller Burnish"
		cause = a
	elseif a == "Sectioning"
		if b == "set up"
			cause = "Changeover"
		else
			cause = "Section"
		end
	elseif a == "Silicon Change"
		cause = "Silicone Change"
	elseif a == "Stop Section"
		cause = "Section"
	elseif a == "Supply Parts"
		if b == "rust"
			cause = "Rusty Parts"
		else
			cause = "No Parts"
		end
	elseif a == "Training"
		cause = "Training"
	elseif a == "Vision System"
		cause = a
	elseif a == "Wash"
		if b == "tooling"
			cause = "Tooling"
		else
			cause = "Wash"
		end
	elseif a == "Weigh 1"
		if b != "tooling"
			cause = "Weigh 1"
		end
	elseif a == "Weigh 2"
		cause = "Weigh 2"
	elseif a == "Weld"
		if b in ["Beam Alignment", "Beam Align"]
			cause = "Beam Alignment"
		elseif b in ["Bias fault", "Bias monitoring fault"]
			cause = "Bias Fault"
		elseif b == "Filament"
			cause = "Filament"
		elseif b in ["section", "Section", "SECTION", "Secton"]
			cause = "Section"
		end
	elseif a == ""
		if b == "Barcodes"
			cause = "Barcode printer"
		elseif b == "manning"
			cause = "Operator Shortage"
		else
			cause = b
		end
	elseif b == "Process"
		cause = a
	elseif b == ""
		cause = a
	end
	cause == "" ? "$a - $b": cause
end

function paint_causes(a, b)
	if a == ""
		if b == ""
			return ""
		end
		if b == " Paint coverage"
			return "Paint Coverage"
		elseif b == "Changeover"
			return b
		elseif uppercase(b) == "COMPRESSOR"
			return "Comrpessor"
		elseif b == "Gaps on line"
			return "Gaps on Line"
		elseif b == "line fill"
			return "Line Fill"
		elseif b == "MEETING"
			return ""
		elseif b == "process"
			return ""
		elseif b == "shortage"
			return "Operator Shortage"
		elseif b in ["Waiting for parts", "Wating for parts"]
			return "No Parts"
		end
		return b
	end

	if a == "Barcode"
		if b == "process"
			return a
		end
	end
	if a == "Bung removal"
		if b == "sequence"
			return a
		end
	end
	if a == "Changeover"
		if b == "line fill"
			return "Line Fill"
		elseif b =="Run Out"
			return b
		end
		return a
	elseif a == "Clean masks"
		return "Clean Masks"
	elseif a == "Clean spindles"
		return "Clean Spindles"
	elseif a == "Clean the bay"
		return "Bay Clean"
	elseif a == "contractors on line"
		return ""
	elseif a == "Fit Paint Mask"
		return a
	elseif a == "Label Printer"
		return a
	elseif a == "Light Guard"
		return a
	elseif uppercase(a) == "LINE FILL"
		if b == "parts jammed"
			return "Parts Jammed"
		end
		if b == "Waiting for parts"
			return "No Parts"
		end
		return "Line Fill"
	elseif a == "Load Spindle"
		return "Line Fill"
	elseif a == "Missing Spindle"
		if b == "Run Out"
			return b
		end
		return "Gaps on Line"
	elseif uppercase(a) == "NO PARTS"
		if b == "Run Out"
			return b
		end
		return "No Parts"
	elseif a == "Oven"
		if b == "parts"
			return "No Parts"
		end
		return "Gaps on Line"
	elseif a == "Packing"
		return a
	elseif a == "Paint Spray"
		if b == "trials"
			return ""
		end
		return "Quality Paint"
	elseif a in ["PPM", "PPMS"]
		return ""
	elseif a == "Quality Checks"
		return a
	elseif a == "Rejects"
		return a
	elseif a == "Ribble parts"
		return ""
	elseif a == "Robot"
		return a
	elseif a == "Run line off"
		return "Run Out"
	elseif a == "Running Slow"
		if b == "Gaps on line"
			return "Gaps on Line"
		elseif b == "Run Out"
			return b
		elseif b =="Quality Paint"
			return b
		elseif b == "hygiene"
			return "Cleaning"
		elseif b == "process"
			return ""
		end
	elseif a == "Service on line"
		return ""
	elseif a in ["T1 Wash", "T2 Rinse", "T4 Phosphate", "T5 Rinse", "T6 tank drop"]
		return titlecase(a)
	elseif uppercase(a) == "TITRATION CHECKS"
		return "Titration Checks"
	elseif a == "Training"
		return a
	elseif a == "Trials"
		return ""
	elseif b == ""
		return a
	end

	return "$a - $b"
end

function write_sheets(io, dfn, line)
	rows = dfn()
	cells = DB_Cols[line]["IncidentNo"]
	if line == "EB"
		println(io, "IncidentNo\tDate\tLeader\tShift\tLine\tProduct\tStd_Rate_PPH\tAvail_Hours\tProduct_Max\tProduct_Actual\tProduct_Variance\tTime_Variance\tEfficiency\tItem\tProcess\tProblem\tLoss_Mins\tEffect\tOEE_Element\tAction\tFix_Or_Repair\tWeld_section_due\tCause")
	else
		println(io, "IncidentNo\tDate\tLeader\tShift\tProduct\tStd_Rate_PPH\tAvail_Hours\tProduct_Max\tProduct_Actual\tProduct_Variance\tTime_Variance\tEfficiency\tLine\tItem\tProcess\tProblem\tQuality_Defect_Type\tQuality_Lost_Parts\tLoss_Mins\tEffect\tOEE_Element\tAction\tFix_Or_Repair\tCause")
	end
	for row in 1:size(rows,1)
		print(io, rows[row, 1])
		print(io, "\t", Dates.format(DateTime(Dates.UTM(rows[row,2])), "yyyy-mm-dd"))
		foreach(c->print(io, "\t", rows[row,c]), 3:cells)
		if line == "EB"
			println(io, "\t", eb_causes(rows[row,15], rows[row,16]))
		else
			println(io, "\t", paint_causes(rows[row,15], rows[row,16]))
		end
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
procEB()
