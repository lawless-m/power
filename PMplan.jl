
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
	wb = Workbook("$home\\power\\PMs\\PM_Plan.xlsx")
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
	write_row!(ws, 0, c, map((d)->Dates.format(d, "yyyy-mm-dd"), sort(collect(keys(events)))))
	
	c = 1
	for d in sort(collect(keys(events)))
		r = 1
		for l in sort(lns)
			lr = r
			if length(events[d][l]) > 0
				for pm in events[d][l]
					lr += write!(ws, lr, c, "$(pms[pm][:ID])\n$(round(pms[pm][:ScheduledHours],1)) Hrs\n$(pms[pm][:Title])", wraptop)
				end
			end
			r += maxs[l]
		end
		c += 1
	end
	
	close(wb)
	
end

write_events(gather_events())

