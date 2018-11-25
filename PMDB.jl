module PMDB

if isfile(raw"c:\users\matt\ntuser.ini")
	push!(LOAD_PATH, raw"C:\Users\matt\Documents\AAM\Lib")
	dbdir = raw"C:\Users\matt\Documents\AAM\SQLite.Data"
	outdir = raw"C:\Users\matt\Documents\AAM"
else
	push!(LOAD_PATH, raw"Z:\Maintenance\Matt-Heath\Lib")
	dbdir = raw"Z:\Maintenance\Matt-Heath\AAM\SQLite.Data"
	outdir = raw"Z:\Maintenance\PPM"
end

using SQLite
using SQLiteTools

export PlexDB, truncate!, column_n, update!, bind!, exebind!, pm_list_by_line, pm_stats_by_date

PlexDB = SQLite.DB("$dbdir\\Plex.db")

inserts = Dict(
			"PM_Stats"=>SQLite.Stmt(PlexDB, "INSERT INTO PM_Stats (Date, Line, OD, High, Items) VALUES(?,?,?,?,?)")
			, "Equipment"=>SQLite.Stmt(PlexDB, "INSERT INTO Equipment (key, ID, Descr, Line) VALUES(?, ?, ?, ?)")
		)

updates = Dict(
			"PMs"=>"UPDATE PM SET LastComplete=?, DueDate=? WHERE ChkKey=? AND Equipment_key=?"
			)
		
bind!(ins::AbstractString, coln::Int, val) = SQLite.bind!(inserts[ins], coln, val)

bind!(ins::AbstractString, vals::Vector, cols::Vector) = bind!(inserts[ins], vals, cols)

exebind!(ins::AbstractString, vals::Vector, cols::Vector) = exebind!(inserts[ins], vals, cols)

update!(up, vals) = SQLite.query(PlexDB, updates[up], values=vals)

pm_list_by_line() = select(PlexDB, "Line, DueDate FROM PM left join Equipment on PM.Equipment_key = Equipment.key WHERE Priority = 'PM' ORDER BY Equipment.Line")

pm_stats_by_date() = table_by(PlexDB, "PM_Stats", "Date")

end