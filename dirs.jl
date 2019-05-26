

if isfile(raw"c:\users\matt\ntuser.ini")
	home = raw"C:\Users\matt\Documents\power"
elseif isfile("/home/matt/.huge")
	home = "/home/matt/"
else
	home = raw"Z:\Maintenance\Matt-H"
end

if ! (joinpath(home, "GitHub", "power" in LOAD_PATH)
	for lib in readdir(joinpath(home, "GitHub"))
		push!(LOAD_PATH, joinpath(home, GitHub, lib))
	end
end

dbdir = joinpath(home, "SQLite.Data")

