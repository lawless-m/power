
for lib in readdir("GitHub")
	push!(LOAD_PATH, realpath("GitHub\\$lib"))
end

if isfile(raw"c:\users\matt\ntuser.ini")
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
else
	dbdir = raw"Z:\Maintenance\Matt-H\power\SQLite.Data"
end

