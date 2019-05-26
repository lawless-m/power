module PAB

include("dirs.jl")

export lineFlds

using ExcelReaders
using PABDB
using SQLiteTools

struct PABwb
    wb
    directory::String
    filename::String
    line::String
    date::Date
    shift::String
    pab
    availability
    performance
    quality
    function PABwb(dir, fn, line, dt, wb, shift, pabws, qualws)
        pab = readxlsheet(wb, pabws, skipstartrows=5, skipstartcols=1, ncols=15, nrows=15)
        avail = readxlsheet(wb, "Availability", skipstartrows=6, skipstartcols=0)
        perf = readxlsheet(wb, "Performance")
        qual = readxlsheet(wb, qualws)
        new(wb, dir, fn, line, dt, shift, pab, avail, perf, qual)
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

for y in 2008:Dates.year(now())+2
    sublist["$y"] = y
end

sublt(a,b) = sublist[a[1]] < sublist[b[1]]

function lineFlds()
    [joinpath(rootdir, d) for d in filter((d)->isdir(joinpath(rootdir, d)), readdir(rootdir))]
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

function shift(fn)
    split(split(fn, '.')[1], '_')[end]
end

function lineDays(ch, dir::String, line::String; since=Date(1970, 1, 1))
    subs = Vector{Tuple{String, String}}()
    for fn in readdir(dir)
        path = joinpath(dir, fn)
        if isdir(path)
            fn[1:3] != "CAM" && push!(subs, (fn, path))
        else
            dbit = datedxlfn(fn)
            if length(dbit) == 1
                dt = txt2date(dbit[1], since)
                if dt > since
                    if ch isa Channel # ugh sorry
                        wb = openxl(joinpath(dir, fn))
                        pabws, qualws = PABsheet(wb), qualitysheet(wb)
                        if pabws == ""
                            println(STDERR, joinpath(dir, fn), " no PAB")
                        end
                        if qualws == ""
                            println(STDERR, joinpath(dir, fn), " no QUAL")
                        end
                        if pabws != "" && qualws != ""
                            put!(ch, PABwb(dir, fn, line, dt, wb, shift(fn), pabws, qualws))
                        end
                    elseif ch isa IOStream
                        println(ch, dir, "\t", fn, "\t", line, "\t", dt, "\t", shift(fn))
                    else
                        println(path, dir, "\t", fn, "\t", line, "\t", dt, "\t", shift(fn))
                    end
                end
            end
        end
    end
    # subdirectories
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
    stage = "Equipment"
    if p.line in ["EB", "HV"]
        if p.availability[2, 3] == "Wash Plant"
            stage = p.availability[2,3]
        end
    end
    c = 3
    while p.availability[3,c] isa String
        stage = p.availability[2,c] isa String ? p.availability[2,c] : stage
        faults[c] = faultID(p.line, stage, p.availability[3,c])
        c += 1
    end
    faults
end

function pabEntry(line, sdte, lastime, part, op, comment, row)

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
    #println(STDERR, "s:", stime, "(", Dates.value(stime), ") e:", etime, "(", Dates.value(etime),  ") l:", line)
    [line, stime, etime, reason, stopt, part, target, op, actual, comment]
end

function availEntry(faults, row)
    ae = Dict{Int, Int}()
    for c in keys(faults)
        if row[c] isa Number
            ae[faults[c]] = floor(Int, row[c])
        end
    end
    ae
end

endtimes = Dict{String, Int}("LATE"=>17, "EARLY"=>11, "NIGHT"=>29, "WEEKEND"=>15)

function PABdata(p::PABwb, io::Union{Void, IOStream}=nothing)
    afaults = availabilityFaults(p)
    endtime = DateTime(p.date) + Dates.Hour(endtimes[uppercase(p.shift)])
    lastime = DateTime(p.date)
    part = ""
    op = ""
    comment = ""
    #for r in 1:size(p.pab, 1)
    #    println("P ", r, " - ", p.pab[r,1:end])
    #end
    #for r in 1:size(p.availability, 1)
    #    println("A ", r, " - ", p.availability[r,1:end])
    #end
    r = 1

    while p.pab[r,1] isa Number && (p.pab[r,4] isa Number || p.pab[r,9] isa Number)

        #for pc in 1:size(p.pab[r],2)
        #    print(pc, ":P>", p.pab[r,pc], "< ")
        #end
        #println()

        pabvals = pabEntry(p.line, p.date, lastime, part, op, comment, p.pab[r, 1:end])
        newpabID = insertPAB!(pabvals, io)
        if newpabID > 0
            for (faultID, loss) in availEntry(afaults, p.availability[r+3, 1:end])
                #println([newpabID, faultID, loss])
                insertAvailLoss!([newpabID, faultID, loss])
            end
        else
            println(STDERR, "?PAB already present")
            #return
        end
        lastime = pabvals[2]
        part = pabvals[6]
        op = pabvals[8]
        comment = pabvals[10]
        #if r == 1
        #    for c = 1:13
        #           print(STDERR, p.pab[r,c], "\t")
        ##    end
        #    println(STDERR, "")
        #end
        r += 1
        #for c = 1:13
        #    print(STDERR, p.pab[r,c], "\t")
        #end
        #println(STDERR, "")
    end
end

function importDailies(since::DateTime=Date(2019, 1, 1))
    io = open(joinpath("c:", "temp", "dailies.sql"), "w+")
    for (d,l) in [("Auto Line", "Auto1"), ("Auto Line 2 Volvo", "Auto2"), ("E.B", "EB"), ("Flexi Line", "Flexi"), ("HV", "HV"), ("Paint Line", "Paint")]
#    for (d,l) in [("Auto Line", "Auto1")]
        for p in Channel(ch->PAB.lineDays(ch, joinpath(PAB.rootdir, d), l, since=since))
            println(p.filename)
            PAB.PABdata(p, io)
        end
    end
end

########
end
