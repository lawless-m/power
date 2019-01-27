
include("../../credentials.jl")
include("dirs.jl")

if isfile(raw"c:\users\matt\ntuser.ini")
	outdir = raw"C:\Users\matt\Documents\power"
else
	outdir = raw"Z:\Maintenance\PPM"
end

using Plexus
using HTMLParser
using Cache
using PMDB
using XlsxWriter

plex = Plex(credentials["plex"]...)

function extract_wrby_year(yr, linecode, lineplx)
	cache("PMS_$(yr)_$(linecode)_csv", ()->work_requests_pms(plex, Date(yr,1,1), Date(yr, 12, 31), true, filter=Plexus.WR_pms, line=lineplx))
end

function get_all_wrs()
	lns = lines()

	for row in 1:size(lns,1)
		for yr in 2018:-1:2013
			extract_pms(yr, lns[1][row], lns[2][row])
		end
	end
end

ymd(d) = Dates.format(d, "yyyy-mm-dd")

overdue(freq, lc, dd) = missInt(lc) + freq - missInt(dd)

function gather_stats()
	stats = Dict{Date, Dict{String, Tuple{Int, Int, Int}}}()
	rows = pm_stats_by_date()
	for row in 1:size(rows,1)
		day = Date(int2time(rows[row, :Date]))
		if ! (day in keys(stats))
			stats[day] = Dict{String, Tuple{Int, Int, Int}}()
		end
		stats[day][rows[row, :Line]] = (rows[row, :OD], rows[row, :High], rows[row, :Items])
	end
	stats
end

function write_stats_sheet(sheets)

	stats = gather_stats()

	(ws, r, c) = sheets["Summary"]
	write_row!(ws, r, 0, ["Date", "Line", "Total OD", "# Tasks", "Max", "Avg"])
	date_entries = sort(collect(keys(stats)))
	for d in date_entries
		for l in sort(collect(keys(stats[d])))
			t, h, n = stats[d][l]
			r = r + 1
			write_row!(ws, r, 0, [ymd(d), l, t,n,h,t/n])
		end
	end

	(ws, r, c) = sheets["Chart"]

	lns = lines()
	r = 1
	for l in lns
		write!(ws, r, 0, l)
		r += 1
	end

	c = 1
	for d in date_entries
		write!(ws, 0, c, ymd(d))
		r = 1
		for l in lns
			write!(ws, r, c, get(stats[d], l, [0])[1])
			r += 1
		end
		c += 1
	end
end

function list_pms()

	wb = Workbook("$outdir\\PM_Tasks.xlsx")
	sheets = Dict{String, Tuple{Worksheet, Int, Int}}()

	wraptop = add_format!(wb, Dict("text_wrap"=>true, "valign"=>"top"))
	boldwraptop = add_format!(wb, Dict("text_wrap"=>true, "valign"=>"top","bold"=>true))
	DueInF = add_format!(wb, Dict("text_wrap"=>true, "valign"=>"top","bold"=>true, "font_color"=>"gray"))
	ODF = add_format!(wb, Dict("text_wrap"=>true, "valign"=>"top","bold"=>true, "font_color"=>"red"))

	function init_sheet!(sht, summary=false)
		sh = add_worksheet!(wb, sht)
		write_row!(sh, 0, 0, ["Line", "Equipment", "Activity", "Priority", "Freq (Days)", "Schd Hours", "TaskList", "Last Complete", "Due Date", "Overdue"], boldwraptop)
		c_widths = [9.0, 35.0, 35.0, 14.0, 6.0, 6.0, 80.0, 12.0, 12.0, 9.0]
		for c in 1:length(c_widths)
			set_column!(sh, c-1, c-1, c_widths[c], wraptop)
		end
		(sh, 1, 0)
	end

	sheets["Summary"] = (add_worksheet!(wb, "Summary"), 0, 0)
	sheets["Chart"] = (add_worksheet!(wb, "Chart"), 0, 0)
	sheets["All"] = init_sheet!("All")

	today = Dates.value(Date(now()))
	totals = Dict{String, Tuple{Int,Int,Int}}()

	for pm in Channel(pm_list)
		for sht in ["All", pm[:Line], pm[:Priority]]
			if ! (sht in keys(sheets))
				sheets[sht] = init_sheet!(sht)
			end
			(ws, r, c) = sheets[sht]

			od = today - Dates.value(pm[:DueDate])
			if od == today
				od = 0
			end
			c += write_row!(ws, r, c, [pm[:Line], pm[:ID], pm[:Title], pm[:Priority], pm[:Frequency], pm[:ScheduledHours], pm[:tasks], ymd(pm[:LastComplete]), ymd(pm[:DueDate])], wraptop)
			if od < 0
				write!(ws, r, c, "Due in $(abs(od))", DueInF)
			else
				if pm[:Priority] == "PM"
					t,n,m = get(totals, sht, (0,0,0))
					totals[sht] = (t+od, n+1, max(m,od))
				end
				write!(ws, r, c, od, ODF)
			end
			r += 1
			sheets[sht] = (ws, r, 0)
		end
	end

	write_stats_sheet(sheets)

	close(wb)
end

function overdues()
	update_pm_dates!((channel)->pm_report(plex, channel))

	today = Dates.value(Date(now()))
	datum = Dates.value(now())

	totals = Dict{String, Tuple{Int,Int,Int}}()

	pms = pm_list_by_line()

	for row in 1:size(pms,1)
		od = today - missInt(pms[row, :DueDate])
		od = od == today ? 0 : od
		if od > 0
			tod,high,items = get(totals, pms[row, :Line], (0,0,0))
			totals[pms[row, :Line]] = (tod+od, max(high,od), items+1)
		end
	end

	foreach((t)->insert_pm_stats!([datum, t[1], t[2][1], t[2][2], t[2][3]]), totals)
end

function event_counts(lns, edays)
	evtcounts = Dict{String, Int}()
	foreach((l)->evtcounts[l]=0, lns)
	events = gather_events(collect(Channel(pm_list)), lns, edays)

	for d in collect(keys(events))
		for ln in lns
			evtcounts[ln] += length(events[d][ln])
		end
	end
	evtcounts
end

function board_stats(prevdays, nextdays)
	lns = lines()
	done_todo = Dict{String, Tuple{Int, Int}}()
	today = Dates.value(Date(now()))
	done = completed(today - prevdays, today)
	todo = event_counts(lns, nextdays)

	stats = gather_stats()
	latest = sort(collect(keys(stats)))[end]
	println(stats[latest])
	println(lns)

	foreach(l->done_todo[l] = (get(done, l, 0), get(todo, l, 0)), lns)
	println("Line\tLast $prevdays\tNext $nextdays\tOD\t#OD")
	foreach(l->println(l, "\t", done_todo[l][1], "\t", done_todo[l][2], "\t", get(stats[latest], l, (0,0,0))[1], "\t", get(stats[latest], l, (0,0,0))[3]), lns)

end

overdues()

list_pms()

board_stats(30, 30)
