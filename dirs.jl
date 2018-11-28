
if isfile(raw"c:\users\matt\ntuser.ini")
	home = raw"C:\Users\matt\Documents\power"
else
	home = raw"Z:\Maintenance\Matt-H\power"	
end

for lib in readdir("$home\\GitHub")
	push!(LOAD_PATH, "$home\\GitHub\\$lib")
end

dbdir = "$home\\SQLite.Data"
