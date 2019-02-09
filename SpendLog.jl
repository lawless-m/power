

include("dirs.jl")

using ExcelReaders
using DataArrays
using XlsxWriter
using Missings
using SQLiteTools


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


function exportSpends()
	wb = Workbook("Z:\\Maintenance\\Matt-H\\Spending\\Collated.xlsx")
	ws = add_worksheet!(wb, "Collated")
	date_format = add_format!(wb, Dict("num_format"=>"dd/mm/yyyy"))
	r = 0
	write_row!(ws, r, 0, ["Date", "Requsitioner", "Supplier", "PA", "PO", "Rcvd", "Goods", "Line", "Reason", "CostCode", "Amount"])
	r += 1
	for row in Channel(extractSpends)
		write!(ws, r, 0, row[1], date_format)
		write_row!(ws, r, 1, row[2:end])
		r += 1
	end
	close(wb)
end

#collateSpends()
exportSpends()
