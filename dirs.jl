
for lib in readdir("GitHub")
	push!(LOAD_PATH, realpath("GitHub\\$lib"))
end
