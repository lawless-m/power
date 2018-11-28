
include("dirs.jl")

using PMDB

pms = collect(Channel(pm_list))

s = Date(now())
e = s + Dates.Day(30)

events = Dict{Date, Vector{Int}}()
for d in s:Dates.Day(1):e
	events[d] = Vector{Int}()
end

for pm in 1:length(pms)
	if pms[pm][:Priority] != "PM"
		continue
	end
	for d in filter((d)->s < d <= e, pms[pm][:LastComplete]:Dates.Day(pms[pm][:Frequency]):e)
		push!(events[d], pm)
	end
end

fid = open("$home\\diary.txt", "w+")
for d in keys(events)
	print(fid, d)
	for pm in events[d]
		print(fid, "\t", pms[pm][:ID], " - ", round(pms[pm][:ScheduledHours],1), " Hrs ", pms[pm][:Title])
	end
	println(fid)
end




