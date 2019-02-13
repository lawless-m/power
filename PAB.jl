module PAB

include("dirs.jl")

export lineFlds

using ExcelReaders
using PABDB

struct PABwb
    wb
    directory::String
    filename::String
    line::String
    date
    pab
    availability
    performance
    quality
    function PABwb(dir, fn, line, dt)
        wb = openxl(dir * "\\" * fn)
        pab = readxlsheet(wb, PABsheet(wb), skipstartrows=5, skipstartcols=1, ncols=15, nrows=15)
        avail = readxlsheet(wb, "Availability", skipstartrows=6, skipstartcols=0)
        perf = readxlsheet(wb, "Performance")
        qual = readxlsheet(wb, qualitysheet(wb))
        new(wb, dir, fn, line, dt, pab, avail, perf, qual)
    end
end

rootdir = "N:\\PAB-OEE Data"
sublist = Dict{String, Int}(
    "January"=>1,
    "February"=>2,
    "March"=>3,
    "April"=>4,
    "May"=>5,
    "June"=>6,
    "July"=>7,
    "August"=>8,
    "September"=>9,
    "October"=>10,
    "october"=>10,
    "November"=>11,
    "Novemeber"=>11,
    "December"=>12,
    "01 - January"=>1,
    "02 - February"=>2,
    "03 - March"=>3,
    "04 - April"=>4,
    "05 - May"=>5,
    "06 - June"=>6,
    "07 - July"=>7,
    "08 - August"=>8,
    "09 - September"=>9,
    "10 - October"=>10,
    "11 - November"=>11,
    "12 - December"=>12,
    "Auto Line 2 Litens USE VOLVO"=>1
)

for y in 208:Dates.year(now())+2
    sublist["$y"] = y
end

sublt(a,b) = sublist[a[1]] < sublist[b[1]]

function lineFlds()
    [rootdir * "\\" * d for d in filter((d)->isdir(rootdir * "\\" * d), readdir(rootdir))]
end

function ltXls(a, b)
    ctime(a) < ctime(b)
end

function datedxlfn(fn)
    isxl() = length(fn) > 5 && fn[1] != '~' && (fn[end-4:end]==".xlsx" || fn[end-3:end]==".xls")
    filter(bits->length(bits) == 10 && bits[end-3:end-1] == "201", isxl() ? split(fn, '_') : [])
end

function txt2date(txt, default)
    try
        return Date(txt, "dd-mm-yyyy")
    end
    try
        return Date(txt, "dd.mm.yyyy")
    end
    println(STDERR, "Not date ", txt)
    default
end

function lineDays(ch, dir::String, line::String; since=Date(1970, 1, 1))
    subs = Vector{Tuple{String, String}}()
    for fn in readdir(dir)
        path = dir * "\\" * fn
        if isdir(path)
            fn[1:3] != "CAM" && push!(subs, (fn, path))
        else
            dbit = datedxlfn(fn)
            if length(dbit) == 1
                dt = txt2date(dbit[1], since)
                if dt > since
                    if ch isa Channel
                        put!(ch, PABwb(dir, fn, line, dt))
                    elseif ch isa IOStream
                        println(ch, dir, "\t", fn, "\t", line, "\t", dt)
                    else
                        println(path, dir, "\t", fn, "\t", line, "\t", dt)
                    end
                end
            end
        end
    end
    foreach(p->lineDays(ch, p[2], line, since=since), sort(subs,lt=sublt,rev=true))
end

function PABsheet(xl)
    for sn in xl.workbook[:sheet_names]()
        if sn[end-2:end] == "PAB"
            return sn
        end
    end
    ""
end

function qualitysheet(xl)
    for sn in xl.workbook[:sheet_names]()
        println(sn)
        if length(sn) > 6 && sn[1:7] == "Quality"
            return sn
        end
    end
    ""
end

function dte(d, f)
    h = floor(Int, f)
    DateTime(d) + Dates.Hour(h) + Dates.Minute(floor(Int, (f-h)*100))
end

function availabilityFaults(p::PABwb)
    faults = Dict{Int, Int}()
    if p.availability[2, 3] != "Wash Plant"
        println(STDERR, "Old Format")
        return
    end
    stage = p.availability[2,3]
    c = 3
    while p.availability[3,c] isa String
        stage = p.availability[2,c] isa String ? p.availability[2,c] : stage
        faults[c] = faultID(p.line, stage, p.availability[3,c])
        c += 1
    end
    faults
end

function pabEntry(line, sdte, lastime, part, op, comment, row)
    if ! (row[1] isa Number)
        return []
    end
    stime = dte(sdte, row[1])
    etime = dte(sdte, row[2])
    if stime < lastime
        stime += Dates.Day(1)
        etime += Dates.Day(1)
    end
    reason = row[3] isa String ? row[3] : ""
    stopt = row[4] isa Number ? floor(Int, row[4]) : 0
    part = length(row[5]) > 2 ? row[5] : part
    target = row[6] isa Number ? floor(Int, row[6]) : 0
    op = length(row[8]) > 1 ? row[8] : op
    actual = row[9] isa Number ? floor(Int, row[9]) : 0
    comment = length(row[end]) > 0 ? (row[end]=="\"" ? comment : row[end] ) : ""
    stime, part, op, comment, [line, Dates.value(stime), Dates.value(etime), reason, stopt, part, target, op, actual, comment]
end

function PABdata(p::PABwb)
    println(p.filename)
    lastime = DateTime(p.date)
    part = ""
    op = ""
    comment = ""
    r = 1
    while p.pab[r,1] isa Number
        lastime, part, op, comment, pabvals = pabEntry(p.line, p.date, lastime, part, op, comment, p.pab[r, 1:end])
        insertPAB(pabvals)
        r += 1
    end
end

###################
end

k = 1
for p in Channel(ch->PAB.lineDays(ch, PAB.rootdir * "\\HV", "HV", since=Date(2019, 2, 1)))
    println(p.filename)
    PAB.PABdata(p)
    println(PAB.availabilityFaults(p))
    k += 1
    if k > 3
        exit()
    end
end
