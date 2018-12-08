
include("dirs.jl")

using PMDB
using XlsxWriter

lns = lines()

function gather_events(pms)
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

function write_events()
	pms = collect(Channel(pm_list))
	events = gather_events(pms)
	
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
		write!(ws, r, c, "Scheduled Hrs:")
		write_formula!(ws, r, c+1, "=sum($range)")
		write!(ws, r+1, c, "Unknown time:")
		write_formula!(ws, r+1, c+1, "=COUNTIF($range,\"=0\")")
		write!(ws, r+2, c, "No. Tasks:")
		write_formula!(ws, r+2, c+1, "=COUNTIF($range,\">=0\")")
		
		c += 2
	end
	
	close(wb)
	
end

#write_events()

function write_for_project()
	wb = Workbook("$home\\power\\PMs\\PM_Proj.xlsx")
	ws = add_worksheet!(wb, "Project")
	date_format = add_format!(wb, Dict("num_format"=>"dd/mm/yyyy hh:mm"))
	
	day(x) = x>0? x/24 : 0
	eod(d) = DateTime(d) + Dates.Minute(1439)
	
	write_row!(ws, 0, 0, ["Deadline", "Resource Names", "Duration", "Name", "Notes"])
	r = 1
	for pm in Channel(pm_list)
		write_datetime!(ws, r, 0, eod(pm[:DueDate]), date_format)
		write_row!(ws, r, 1, [pm[:ID], day(pm[:ScheduledHours]), pm[:Title], pm[:Priority]])
		r += 1
	end
	close(wb)
end

write_for_project()


