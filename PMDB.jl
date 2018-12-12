module PMDB

include("dirs.jl")

using SQLite
using SQLiteTools
using DataArrays

export PlexDB, pm_list_by_line, pm_stats_by_date, fill_equipment, lines, insert_pm, clear_pm_tasks, clear_pms, pm_list, update_pm_dates!, insert_pm_stats!

export missInt, int2date, int2time

PlexDB = SQLite.DB("$dbdir\\Plex.db")

missInt = SQLiteTools.missInt
int2date = SQLiteTools.int2date
int2time = SQLiteTools.int2time

inserts = Dict()

pm_list_by_line() = SQLite.query(PlexDB, "SELECT Line, DueDate FROM PM left join Equipment on PM.Equipment_key = Equipment.key WHERE Priority = 'PM' ORDER BY Equipment.Line")

pm_stats_by_date() = table_by(PlexDB, "PM_Stats", "Date")

function fill_equipment(eqs)
	truncate!(PlexDB, "Equipment")
	ins = SQLite.Stmt(PlexDB, "INSERT INTO Equipment (key, ID, Descr, Line) VALUES(?, ?, ?, ?)")
	for (k,id,line) in eqs
		exebind!(ins, [k, id, id, line])
	end
end

clear_pm_tasks() = truncate!(PlexDB, "PM_task")
clear_pms() = truncate!(PlexDB, "PM")

function insert_pm(ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate, tasks)
	exebind!(get!(inserts, "PM", SQLite.Stmt(PlexDB, "INSERT INTO PM (ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate) VALUES(?,?,?,?,?,?,?,?,?)")), [ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate])
		
	foreach((t)->exebind!(get!(inserts, "PM_Task", SQLite.Stmt(PlexDB, "INSERT INTO PM_Task (Task, Instructions, ChkNo, Equipment_key) VALUES(?,?,?,?)")), [t, "", ChkNo, Equipment_key]), tasks)
end

lines() = SQLite.query(PlexDB, "Select name FROM Lines")[1:end, 1]

function pm_list(chan)
	pms = SQLite.query(PlexDB, "select PM.ChkNo, PM.ChkKey, PM.Title, PM.Priority, PM.Frequency, PM.ScheduledHours, PM.LastComplete, PM.DueDate, PM.Equipment_key, Equipment.ID, Equipment.Line from PM left join Equipment on PM.Equipment_key = Equipment.key and Equipment.key is not null ORDER BY Equipment.Line, PM.Priority, PM.DueDate")
	for row in 1:size(pms,1)
		if typeof(pms[row, :ID]) == Missings.Missing
			continue
		end
		
		pm = Dict{Symbol,Any}()
		foreach((s)->pm[s] = pms[row, s], [:Line, :ID, :Title, :Priority, :Frequency, :ScheduledHours])
		tasks = pm_tasks(pms[row, :Equipment_key], pms[row, :ChkNo])
		pm[:tasks] = ""
		for trow in 1:size(tasks,1)
			if trow > 1
				pm[:tasks] = pm[:tasks] * "\n"
			end
			pm[:tasks] = pm[:tasks] * " * " * tasks[trow, :Task]
			if tasks[trow, :Instructions] != ""
				pm[:tasks] = pm[:tasks] * " [" * tasks[trow, :Instructions] * "]"
			end
		end
		pm[:LastComplete] = int2date(missInt(pms[row, :LastComplete]))
		pm[:DueDate] = int2date(missInt(pms[row, :DueDate]))
		put!(chan, pm)
	end
end

function pm_tasks(Equipment_key, ChkNo)
	SQLite.query(PlexDB, "SELECT Task, Instructions from PM_Task WHERE Equipment_key=? AND ChkNo=?", values=[Equipment_key, ChkNo])
end

update_pm_dates!(src) = foreach((vals)->SQLite.query(PlexDB, "UPDATE PM SET LastComplete=?, DueDate=? WHERE ChkKey=? AND Equipment_key=?", values=vals), Channel(src))	

insert_pm_stats!(vals) = exebind!(get!(inserts, "PM_Stats", SQLite.Stmt(PlexDB, "INSERT INTO PM_Stats (Date, Line, OD, High, Items) VALUES(?,?,?,?,?)")), vals)

end
