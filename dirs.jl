

Home = ""

if isfile(raw"c:\users\matt\ntuser.ini")
	Home = raw"C:\Users\matt\Documents\power"
elseif isfile("/home/matt/.huge")
	Home = "/home/matt/"
else
	Home = raw"Z:\Maintenance\Matt-H"
end

if ! (joinpath(Home, "GitHub", "power") in LOAD_PATH)
	for lib in readdir(joinpath(Home, "GitHub"))
		push!(LOAD_PATH, joinpath(Home, "GitHub", lib))
	end
end

DBDir = joinpath(Home, "SQLite.Data")
