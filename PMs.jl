
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
using DataArrays
using SQLite


plex = Plex(credentials["plex"]...)


function lines()
	SQLite.query(PlexDB, "select * from Lines")[1:end]
end


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


missInt(v) = typeof(v) == Missings.Missing ? 0 : v

int2date(dv) = dv > 0 ? ymd(Date(Dates.UTD(missInt(dv)))) : ""

int2time(dv) = dv > 0 ? DateTime(Dates.UTM(missInt(dv))): DateTime(now())

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
	lines = column_n("Lines", 1)
	r = 1
	for l in lines
		write!(ws, r, 0, l)
		r += 1
	end 
	
	c = 1
	for d in date_entries
		write!(ws, 0, c, ymd(d))
		r = 1
		for l in lines
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
	
	
	pms = SQLite.query(PlexDB, "select PM.ChkNo, PM.ChkKey, PM.Title, PM.Priority, PM.Frequency, PM.ScheduledHours, PM.LastComplete, PM.DueDate, PM.Equipment_key, Equipment.ID, Equipment.Line from PM left join Equipment on PM.Equipment_key = Equipment.key ORDER BY Equipment.Line, PM.Priority, PM.DueDate")
	
	today = Dates.value(Date(now()))
	
	totals = Dict{String, Tuple{Int,Int,Int}}()
		
	for row in 1:size(pms,1)			
		tasks = SQLite.query(PlexDB, "SELECT Task, Instructions from PM_Task WHERE Equipment_key=? AND ChkNo=?", values=[pms[row, :Equipment_key], pms[row, :ChkNo]])
		
		tsks = ""
		for trow in 1:size(tasks,1)
			if trow > 1
				tsks = tsks * "\n"
			end
			tsks = tsks * " * " * tasks[trow, :Task]
			if tasks[trow, :Instructions] != ""
				tsks = tsks * " [" * tasks[trow, :Instructions] * "]"
			end
		end
		
		for sht in ["All", pms[row, :Line], pms[row, :Priority]]
			if ! (sht in keys(sheets))
				sheets[sht] = init_sheet!(sht)
			end
			(ws, r, c) = sheets[sht]
			
			od = today - missInt(pms[row, :DueDate])
			if od == today
				od = 0
			end
			
			c += write_row!(ws, r, c, [
				pms[row, :Line], 
				pms[row, :ID], 
				pms[row, :Title],
				pms[row, :Priority],
				pms[row, :Frequency],
				pms[row, :ScheduledHours],	
				tsks,
				int2date(missInt(pms[row, :LastComplete])),
				int2date(missInt(pms[row, :DueDate]))
			], wraptop)
			if od < 0
				write!(ws, r, c, "Due in $(abs(od))", DueInF)	
			else
				if pms[row, :Priority] == "PM"
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

	foreach((p)->update!("PMs", p), pm_report(plex))
	
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
	
	bind!("PM_Stats", 1, datum)
	foreach((t)->exebind!("PM_Stats", [t[1], t[2][1], t[2][2], t[2][3]], [2,3,4,5]), totals)
end


#pm2sqlite()

overdues()

list_pms()






