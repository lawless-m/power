
if isfile(raw"c:\users\matt\ntuser.ini")
	home = raw"C:\Users\matt\Documents\power"
else
	home = raw"Z:\Maintenance\Matt-H"
end

if ! ("$home\\GitHub\\power" in LOAD_PATH)
	for lib in readdir("$home\\GitHub")
		push!(LOAD_PATH, "$home\\GitHub\\$lib")
	end
end

dbdir = "$home\\SQLite.Data"
