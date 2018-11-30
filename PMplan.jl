
include("dirs.jl")

using PMDB
using XlsxWriter

pms = collect(Channel(pm_list))
lns = lines()

function gather_events()
	s = Date(now())
	e = s + Dates.Day(30)
	events = Dict{Date, Dict{String, Vector{Int}}}()
	for d in s:Dates.Day(1):e
		events[d] = Dict{String, Vector{Int}}()
		for l in lns
			events[d][l] = Vector{Int}()
		end
	end

	for pm in 1:length(pms)
		if pms[pm][:Priority] != "PM"
			continue
		end
		for d in filter((d)->s < d <= e, pms[pm][:LastComplete]:Dates.Day(pms[pm][:Frequency]):e)
			push!(events[d][pms[pm][:Line]], pm)
		end
	end
	events
end

function max_items_per_line(events)
	maxs = Dict{String, Int}()
	foreach((l)->maxs[l] = 0, lns)
	for d in keys(events)
		for l in lns
			maxs[l] = max(maxs[l], length(events[d][l]))
		end
	end
	maxs
end

function write_events(events)
	wb = Workbook("$home\\power\\PMs\\PM_Plan2.xlsx")
	ws = add_worksheet!(wb, "Plan")
	wraptop = add_format!(wb, Dict("text_wrap"=>true, "valign"=>"top"))

	maxs = max_items_per_line(events)
	
	r = 1
	for l in sort(collect(keys(maxs)))
		if maxs[l] == 0
			continue
		end
		write!(ws, r, 0, l, wraptop)
		r += maxs[l]
	end
	c = 1
	for d in sort(collect(keys(events)))
		c += 1 + write!(ws, 0, c, Dates.format(d, "yyyy-mm-dd"))
	end
	
	c = 1
	for d in sort(collect(keys(events)))
		r = 1
		for l in sort(lns)
			ltxt = " - $l"
			lr = r
			if length(events[d][l]) > 0
				for pm in events[d][l]
					write!(ws, lr, c, round(pms[pm][:ScheduledHours],1), wraptop)
					id = endswith(pms[pm][:ID], ltxt) ? pms[pm][:ID][1:end-length(ltxt)] : pms[pm][:ID]
					lr += write!(ws, lr, c+1, "$(id)\n$(pms[pm][:Title])", wraptop)
				end
			end
			r += maxs[l]
		end
		range = rc2cell(1, c) * ":" * rc2cell(r-1, c)
		write_formula!(ws, r, c, "=concatenate(\"Scheduled Hrs: \", sum($range))")
		write_formula!(ws, r+1, c, "=concatenate(\"Unknown time: \", COUNTIF($range,\"=0\"))")
		write_formula!(ws, r+2, c, "=concatenate(\"No. Tasks: \", COUNTIF($range,\">=0\"))")
		
		c += 2
	end
	
	close(wb)
	
end

write_events(gather_events())

