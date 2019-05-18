

include("dirs.jl")

using ExcelReaders
using DataArrays
using XlsxWriter
using Missings
using SQLiteTools

codes = Dict{String,String}()
for (t,c) in Dict("Buildings 112"=>"661534", "Lifting Equipment"=>"660428", "Consumables"=>"660400", "Lubricants"=>"660415", "Machine Repairs"=>"660417", "Safety Equip"=>"660429", "Fork Truck Maint"=>"660430", "Maint Contracts"=>"660456", "Building Repairs"=>"661520", "Misc Plant"=>"661529", "Hire of Plant & M/c"=>"662645", "Spares"=>"100001008 US GAAP", "EnviCare & Maint"=>"661533", "Buildings"=>"661520", "Capital"=>"---", "MRO Spares"=>"660418-311", "BMW"=>"5000 19607", "Storage Units"=>"660416-115", "Software"=>"662643-144")
	codes[c] = t
end

xl = openxl("Z:\\Maintenance\\Monthly Spending Budget\\Monthly Spend Log.xlsx")
SpendDB = SQLite.DB("$dbdir\\Spending.db")

function MMMn(t)
	if length(t) >= 3
		findfirst(isequal(t[1:3]), ["Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec"])
	else
		0
	end
end

function month_tabs(ch)
	for s in filter((s) -> ismatch(r"^[a-zA-Z]+[ ]*[0-9]+", s), xl.workbook[:sheet_names]())
		m, y = match(r"^([a-zA-Z]+)[ ]*([0-9]+)", s).captures
		if MMMn(m) > 0
			push!(ch, s)
		end
	end
end

str(t::DataValues.DataValue{Union{}}) = ""
str(t::Float64) = string(convert(Int64,round(t,0)))
str(t::Int) = string(convert(Int64,t))
str(t::String) = t
flt(t::Float64) = t
flt(t::DataValues.DataValue{Union{}}) = 0.0

function spends(tab)
	d = readxlsheet(xl, tab, skipstartrows=4)
	hdrs = map(str, d[1,1:end])
	ins = SQLite.Stmt(SpendDB, "INSERT INTO Spend (Date, Requsitioner, Supplier, PA, PO, Rcvd, Goods, Line, Reason, CostCode, Amount) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
	rows,cols = size(d)
	r = 2
	while r < rows && typeof(d[r,1]) != DataValues.DataValue{Union{}}
		dv = 0
		try
			dv = Dates.value(Date(d[r,1]))
		catch y
			println("Bad Date on $tab ", d[r,1])
		end

		for c in 10:length(hdrs)
			if isa(d[r,c], Number) && d[r,c] > 0
				exebind!(ins, [dv, str(d[r,2]), str(d[r,3]), str(d[r,4]), str(d[r,5]), d[r,6]==""?0:1, str(d[r,7]), str(d[r,8]), str(d[r,9]), hdrs[c], flt(d[r,c])])
			end
		end
		r = r + 1
	end
end

function extractSpends(ch)
	d = SQLite.query(SpendDB, "SELECT Date, Requsitioner, Supplier, PA, PO, Rcvd, Goods, Line, Reason, CostCode, Amount FROM Spend")
	foreach(r->put!(ch, [int2date(d[r,1]), d[r,2], d[r,3], d[r,4], d[r,5], d[r,6]==1?"x":"", d[r,7], d[r,8], d[r,9], d[r,10], d[r,11]]), 1:size(d,1))
end

function collateSpends()
	truncate!(SpendDB, "Spend")
	foreach(spends, Channel(month_tabs))
end


function exportSpends(wb, df)
	ws = add_worksheet!(wb, "Collated")
	r = 0
	write_row!(ws, r, 0, ["Date", "Requsitioner", "Supplier", "PA", "PO", "Rcvd", "Goods", "Line", "Reason", "CostCode", "Amount"])
	r += 1
	for row in Channel(extractSpends)
		write!(ws, r, 0, row[1], date_format)
		write_row!(ws, r, 1, row[2:end])
		r += 1
	end
end

function spendGrid(wb)
	d = SQLite.query(SpendDB, "SELECT distinct CostCode FROM Spend")
	ws = add_worksheet!(wb, "Grid")
	cc_col = Dict{String, Int}()
	for c in 1:size(d,1)
		write!(ws, 0, c, codes[d[c,1]])
		cc_col[d[c,1]] = c
	end

	yq_row = Dict{String, Int}()
	r = 1
	for y in 2010:2019
		for q = 1:4
			yq_row["$y-Q$q"] = r
			write!(ws, r, 0, "$y-Q$q")
			r += 1
		end
	end

	spend = Dict{String, Dict{String, Float64}}()
	d = SQLite.query(SpendDB, "SELECT Date, Requsitioner, Supplier, PA, PO, Rcvd, Goods, Line, Reason, CostCode, Amount FROM Spend")
	for r in 1:size(d,1)
		dt = int2date(d[r,1])
		yq = "$(Dates.year(dt))-Q$(Dates.quarterofyear(dt))"
		cc = get!(spend, yq, Dict{String, Float64}())
		a = get!(cc, d[r,10], 0.0)
		spend[yq][d[r,10]] = a + d[r,11]
	end

	for (yq, cc) in spend, (c,a) in cc
		try
			write!(ws, yq_row[yq], cc_col[c], a)
		end
	end
end

collateSpends()
wb = Workbook("Z:\\Maintenance\\Matt-H\\Spending\\Collated.xlsx")
date_format = add_format!(wb, Dict("num_format"=>"dd/mm/yyyy"))
exportSpends(wb, date_format)
spendGrid(wb)
close(wb)
