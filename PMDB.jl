module PMDB

if isfile(raw"c:\users\matt\ntuser.ini")
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
else
	dbdir = raw"Z:\Maintenance\Matt-H\power\SQLite.Data"
end

using SQLite
using SQLiteTools

export PlexDB, pm_list_by_line, pm_stats_by_date, fill_equipment, lines, insert_pm, clear_pm_tasks, clear_pms

PlexDB = SQLite.DB("$dbdir\\Plex.db")

inserts = Dict(
			"PM_Stats"=>SQLite.Stmt(PlexDB, "INSERT INTO PM_Stats (Date, Line, OD, High, Items) VALUES(?,?,?,?,?)")
			, "Equipment"=>SQLite.Stmt(PlexDB, "INSERT INTO Equipment (key, ID, Descr, Line) VALUES(?, ?, ?, ?)")
			, "PM"=>SQLite.Stmt(PlexDB, "INSERT INTO PM (ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate) VALUES(?,?,?,?,?,?,?,?,?)")
			, "PM_Task"=>SQLite.Stmt(PlexDB, "INSERT INTO PM_Task (Task, Instructions, ChkNo, Equipment_key) VALUES(?,?,?,?)")	
		)

updates = Dict(
			"PMs"=>"UPDATE PM SET LastComplete=?, DueDate=? WHERE ChkKey=? AND Equipment_key=?"
			)
			
update!(up, vals) = SQLite.query(PlexDB, updates[up], values=vals)

pm_list_by_line() = SQLite.query(PlexDB, "SELECT Line, DueDate FROM PM left join Equipment on PM.Equipment_key = Equipment.key WHERE Priority = 'PM' ORDER BY Equipment.Line")

pm_stats_by_date() = table_by(PlexDB, "PM_Stats", "Date")

function fill_equipment(eqs)
	truncate!(PlexDB, "Equipment")
	for (k,id,line) in eqs
		exebind!(inserts["Equipment"], [k, id, id, line])
	end
end

clear_pm_tasks() = truncate!(PlexDB, "PM_task")
clear_pms() = truncate!(PlexDB, "PM")

function insert_pm(ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate, tasks)
	exebind!(inserts["PM"], [ChkNo, ChkKey, Equipment_key, Title, Priority, Frequency, Instructions, ScheduledHours, StartDate])
	for t in tasks
		exebind!(inserts["PM_Task"], [t, "", ChkNo, Equipment_key])
	end
end

lines() = column_n(PlexDB, "Lines", 1)

end

