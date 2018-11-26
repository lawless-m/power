module EBDB

using SQLite

if isfile(raw"c:\users\matt\ntuser.ini")
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
else
	dbdir = raw"Z:\Maintenance\Matt-H\power\SQLite.Data"
end

EBdb = SQLite.DB("$dbdir\EB_Perf.db")


end
