
include("dirs.jl")

using ExcelReaders



xl = openxl("C:\\Users\\heathm\\Desktop\\TEMP\\build plan.xlsx")

lines = Dict("EB"=>"EB",
    "Dieburg"=>"LVC",
    "CAM LINE"=>"Auto 1", "Cam Line"=>"Auto 1", "cam line"=>"Auto 1", "CAM"=>"Auto 1", "Cam line"=>"Auto 1",
    "paintline"=>"Paint", "Paintline"=>"Paint", "Paint line"=>"Paint",
    "AL2"=>"Auto 2", "Auto Line 2"=>"Auto 2", "AL 2"=>"Auto 2", "Auto 2"=>"Auto 2",
    "line 5"=>"Flexi", "Line 5"=>"Flexi",
    "HV"=>"HV", "H.V"=>"HV")


function line(sname)
    for (k,v) in lines
        if contains(sname, k) return v end
    end
    println(sname)
    return "??"
end

line_parts = Dict{String, Dict}()

for sht in xl.workbook["sheet_names"]()
    data = readxlsheet(xl, "$(sht)")
    l = line(sht)
    prts = get!(line_parts, l, Dict{String, Set{String}}())
    for r in 3:size(data, 1)
        if typeof(data[r,2]) == String
            ds = get!(prts, data[r,2], Set{String}())
            push!(ds, typeof(data[r,3])==String ? strip(data[r,3]) : "")
            prts[data[r,2]] = ds
        end
    end
end

fid = open("c:\\temp\\parts.csv", "w+")
for (l, ps) in line_parts
    for (p, ds) in ps
        for d in ds
            if d != ""
                println(fid, l, "\t", p, "\t", d)
            end
        end
    end
end
close(fid)
