module PMDB

if isfile(raw"c:\users\matt\ntuser.ini")
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
else
	dbdir = raw"Z:\Maintenance\Matt-H\power\SQLite.Data"
end

using SQLite
using SQLiteTools

export PlexDB, truncate!, column_n, update!, bind!, exebind!, pm_list_by_line, pm_stats_by_date, column_n

PlexDB = SQLite.DB("$dbdir\\Plex.db")

inserts = Dict(
			"PM_Stats"=>SQLite.Stmt(PlexDB, "INSERT INTO PM_Stats (Date, Line, OD, High, Items) VALUES(?,?,?,?,?)")
			, "Equipment"=>SQLite.Stmt(PlexDB, "INSERT INTO Equipment (key, ID, Descr, Line) VALUES(?, ?, ?, ?)")
		)

updates = Dict(
			"PMs"=>"UPDATE PM SET LastComplete=?, DueDate=? WHERE ChkKey=? AND Equipment_key=?"
			)
			
update!(up, vals) = SQLite.query(PlexDB, updates[up], values=vals)

pm_list_by_line() = SQLite.query(PlexDB, "SELECT Line, DueDate FROM PM left join Equipment on PM.Equipment_key = Equipment.key WHERE Priority = 'PM' ORDER BY Equipment.Line")

pm_stats_by_date() = table_by(PlexDB, "PM_Stats", "Date")

column_n(table::String, col::Integer) = SQLiteTools.column_n(PlexDB, table, col)

end