module PABOEE

export Event, unscheduled, downPC, slot_txt

using ExcelReaders
using DataArrays

struct Event
	line::String
	slot_start::DateTime
	slot_end::DateTime
	downtime::Base.Dates.Minute
	equipment::String
end

slot_txt(e::Event) = Base.Dates.format(e.slot_start, "HH:MM")"-"Base.Dates.format(e.slot_end, "HH:MM")

function downPC(t::Base.Dates.Minute, ts::Base.Dates.DateTime, te::Base.Dates.DateTime)
	# downtime as a % of the interval
	ms = Base.Dates.datetime2epochms(te) - Base.Dates.datetime2epochms(ts)
	s = 60convert(Int64, Dates.value(t))
	#println(ts, " - ", te, " => ", 100000s, " / ", ms)
	round(1000s / ms, 3)
end

function show(io, e::Event)
	println(io, "\t", e.line, " ", slot_txt(e), " ", trunc(Int64, 100downPC(e.downtime, e.slot_start, e.slot_end)), "% lost on ", e.equipment)
end


struct XL
	dir::String
	line::String
	fn::String
	date::Date
end

lines(dir) = [d for d in readdir(dir) if !(length(d) > 4 && d[1:4] == "New ")]
add_xl!(xls, dir, line, fn, dte) = push!(xls, XL(dir, line, fn, dte))
fn_date(fn) = Date(split(fn, "_")[3], "dd-mm-yyyy")

function maybe_add_xlfn!(xls, dir, line, fn, since)
	if length(fn) > 5 && fn[1] != '~' && fn[end-4:end] in (".xlsx", ".xltm")
		println(STDERR, fn)
		dte = fn_date(fn)
		since <= dte <= (since + Base.Dates.Month(1)) && add_xl!(xls, dir, line, fn, dte)
	end 
end

add_root_xldir!(xls, dir, line, since) = foreach((fn)->maybe_add_xlfn!(xls, "$dir/$line", line, fn, since), readdir("$dir/$line"))

function add_xl_month!(xls, dir, line, month)
	mm = Base.Dates.monthname(month)
	yy = Base.Dates.value(Base.Dates.Year(month))
	isdir("$dir/$yy/$mm") && foreach((fn)->maybe_add_xlfn!(xls, "$dir/$yy/$mm", line, fn, month), readdir("$dir/$yy/$mm"))
end

function find_xls!(xls, root, month)
	for line in lines(root)
		add_root_xldir!(xls, root, line, month)
		add_xl_month!(xls, "$root/$line", line, month)
	end
end


function isnotnumeric(el)
	typeof(el) == Float64 && (return false)
	typeof(el) == String && ismatch(r"^[0-9]+\.?[0-9]*", el) && (return false)
	true
end


function numerical(el)
	typeof(el) == Float64 && (return el)
	typeof(el) == String && ismatch(r"^[0-9]+\.?[0-9]*", el) && (return parse(el))
	0
end

function totalrow(xl)
	for r in 1:size(xl,1)
		if typeof(xl[r,1]) == String && uppercase(xl[r,1]) == "TOTAL"
			r -= 1
			while isnotnumeric(xl[r,1])
				r -=1
			end
			return r
		end
	end
end

function timerows_equipment(xl)
	equip = Vector{String}()
	endr = totalrow(xl)
	equipr = 1 + endr - findfirst(isnotnumeric, reverse(xl[1:endr, 1]))
	
	c = 3
	while typeof(xl[equipr,c]) == String
		push!(equip, xl[equipr,c])
		c += 1
	end
	return (equipr+1):endr, equip
end

function time(d::Date, f::Float64)
	h = trunc(Int64, f)
	m = trunc(Int64, 100round(f-h, 2))
	d + Base.Dates.Time(mod(h, 24), m, 0)
end

time(d::Date, f::Int64) = d + Base.Dates.Time(mod(f, 24), 0, 0)
fill_unsched!(events, xls) = foreach((xl)->add_unscheduled!(events, xl), xls)

function add_unscheduled!(events, xl)
	println(STDERR, xl.fn)
	sheet = readxlsheet("$(xl.dir)/$(xl.fn)", "Availability")
	se_rows, equip = timerows_equipment(sheet)
	midnight = xl.date
	for t in se_rows, e in 3:3+length(equip)
		lost = numerical(sheet[t,e])
		if lost > 0
			ts = time(midnight, numerical(sheet[t,1]))
			te = time(midnight, numerical(sheet[t,2]))
			ts > te && (te += Base.Dates.Hour(24))
			push!(get!(events, Date(ts), Vector{Event}()), Event(xl.line, ts, te, Base.Dates.Minute(lost), equip[e-2]))
			Date(te) > midnight && (midnight += Base.Dates.Day(1))
		end
	end
end

root = "N:/PAB-OEE Data"


function unscheduled(first::DateTime)
	xls = Vector{XL}()
	find_xls!(xls, root, first)
println(xls)
exit()
	eventsbd = Dict{Date, Vector{Event}}()
	fill_unsched!(eventsbd, xls)
	events = Vector{Event}()
	for day in sort(collect(keys(eventsbd)))
		foreach((e)->push!(events, e), sort(eventsbd[day], lt=(a,b)->a.slot_start<b.slot_start))
	end
	events
end


function test()
	fn = "Auto2 Volvo_OEE_25-08-2018_Early.xlsx"
	line = "Auto Line 2 Volvo"
	dir = "N:/PAB-OEE Data/Auto Line 2 Volvo"
	day = DateTime(2018, 8, 25)
	eventsbd = Dict{Date, Vector{Event}}()
	
	eventsbd = Dict{Date, Vector{Event}}()
	fill_unsched!(eventsbd, [XL(dir, line, fn, day)])
end

test()

end