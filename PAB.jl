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

function shift(fn)
    split(split(fn, '.')[1], '_')[end]
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
                    if ch isa Channel # ugh sorry
                        wb = openxl(dir * "\\" * fn)
                        pabws, qualws = PABsheet(wb), qualitysheet(wb)
                        if pabws == ""
                            println(STDERR, dir, "\\", fn, " no PAB")
                        end
                        if qualws == ""
                            println(STDERR, dir, "\\", fn, " no QUAL")
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

###################
end

using SQLiteTools
using XlsxWriter

function importDailies(since::DateTime=Date(2019, 1, 1))
    io = open("c:\\temp\\dailies.sql", "w+")
    for (d,l) in [("Auto Line", "Auto1"), ("Auto Line 2 Volvo", "Auto2"), ("E.B", "EB"), ("Flexi Line", "Flexi"), ("HV", "HV"), ("Paint Line", "Paint")]
#    for (d,l) in [("Auto Line", "Auto1")]
        for p in Channel(ch->PAB.lineDays(ch, PAB.rootdir * "\\" * d, l, since=since))
            println(p.filename)
            PAB.PABdata(p, io)
        end
    end
end

struct Wb
    wb
    fmts::Dict
    function Wb()
        wb = Workbook("Z:\\Maintenance\\Matt-H\\power\\OEE\\Avail.xlsx")
        fmts = Dict()
        fmts["angle"] = add_format!(wb, Dict("rotation"=>45))
        fmts["day_fmt"] = add_format!(wb, Dict("num_format"=>"dd/mm/yy"))
        fmts["time_fmt"] = add_format!(wb, Dict("num_format"=>"hh:mm"))
        new(wb, fmts)
    end
end

function availabilityColour(startt, endt)
    wb = Wb()
    for line in ["Auto1", "Auto2", "EB", "Flexi", "HV", "Paint"]
        availabilityColourLine(wb, line, startt, endt)
    end
    close(wb.wb)
end


function availabilityColourLine(wb::Wb, line, startt, endt)
    ws = add_worksheet!(wb.wb, line)
    flts = PABDB.faultList()
    fault_cols = Dict{Int, Int}()

    write!(ws, 1, 0, "Day")
    set_column!(ws, 1, 0, 10)
    write!(ws, 1, 1, "Start")
    set_column!(ws, 1, 1, 5)
    write!(ws, 1, 2, "End")
    set_column!(ws, 1, 2, 5)

    stagec = 3

    for s in keys(flts[line])
        write!(ws, 0, stagec, s, wb.fmts["angle"])
        for e in keys(flts[line][s])
            write!(ws, 1, stagec, e, wb.fmts["angle"])
            set_column!(ws, 1, stagec, 3)
            fault_cols[flts[line][s][e]] = stagec
            stagec += 1
        end
    end

    pabs = PABDB.pabsBetween(startt, endt)
    aloss = PABDB.availLossBetween(startt, endt)

    outr = 2
    for r in 1:size(pabs,1)
        if pabs[r, :Line] == line
            write!(ws, outr, 0, int2time(pabs[r, :StartT]), wb.fmts["day_fmt"])
            write!(ws, outr, 1, int2time(pabs[r, :StartT]), wb.fmts["time_fmt"])
            write!(ws, outr, 2, int2time(pabs[r, :EndT]), wb.fmts["time_fmt"])

            for a in 1:size(aloss,1)
                if aloss[a, :PAB_ID] == pabs[r, :id]
                    write!(ws, outr, fault_cols[aloss[a, :Fault_ID]], aloss[a, :Loss])
                end
            end

            write!(ws, outr, stagec, pabs[r, :Comment])
            outr += 1
        end
    end
end

function MTBF(line, st, se)
    faults = PABDB.idfaults(line)

    println(faults[35])

    availtime = Dict{Int, Vector{Int}}()
    downtime = Dict{Int, Vector{Tuple{DateTime, DateTime, Int, Int}}}()
    for k in keys(faults)
        downtime[k] = Vector{Int}()
    end
    lstart = 0

    slots = PABDB.all_slots(line, st, se) #  line, startT, endT, stopmins, loss, fault_id

    if size(slots, 1) == 0
        return
    end
    earliest = int2time(slots[1, :startT])
    latest = int2time(slots[end, :startT])
    for r in 1:size(slots, 1)
        #println(slots[r, 1:end])
        if typeof(slots[r, :loss]) != Missings.Missing
            push!(downtime[slots[r,:fault_id]], (int2time(slots[r, :startT]), int2time(slots[r, :endT]), 60 - slots[r, :stopmins], slots[r, :loss]))
        end
    end
    contiguous = Dict{Int, Vector{Int}}()
    tbf = Dict{Int, Vector{Int}}()
    started = earliest
    ended = earliest
    for f in keys(faults)
        contiguous[f] = Vector{Int}()
        tbf[f] = Vector{Int}()

        loss = 0

        for d in downtime[f] # d = (startt, endt, avail, loss)

            println("$f - $d - $loss")

            if d[3] == d[4] # whole of this period was lost
                println("$f p1 $loss")
                ended = min(ended, d[1])
                if ended == d[1] # this event ended the timer
                    println("$f p1a $loss")
                    push!(tbf[f], Dates.value(Dates.Minute(ended - started)))
                else

                    println("$f p1b $loss")
                    # this event was over two contiguous periods and already ended
                end
                started = d[2]
                loss += d[4]
                continue
            end

            # last event ended when this one started so put it at the start of the period, add this loss on to the end then restart timer
            if d[1] == started
                println("$f p2 $loss")
                loss += d[4]
                push!(contiguous[f], loss)
                loss = 0
                started += Dates.Minute(d[4])
                continue
            end

            if loss > 0 # previous loss was whole period and not contiguous with this one
                println("$f p3 $loss")
                started = d[1]
                ended = d[2]
                push!(contiguous[f], loss)
                loss = 0
                continue
            end

            println("$f p4 $loss")
            # this is a new event, put it at the end of the period
            ended = d[2] - Dates.Minute(d[4])
            push!(tbf[f], Dates.value(Dates.Minute(latest - started)))
            started = d[2]
            ended = d[2]
            loss = d[4]

        end

        push!(tbf[f], Dates.value(Dates.Minute(ended - started)))
        if loss > 0 # previous event was the last
            push!(contiguous[f], loss)
            loss = 0
            started = earliest
            ended = earliest
        end

    end
    println(line, "\n From $st to $se")
    for k in keys(downtime)
        avail = sum(tbf[k]) - sum(contiguous[k])
        if avail != 0
            println("avail $avail, tbf ", tbf[k])
        end
        if length(contiguous[k]) > 0
            println(faults[k][1], " - ", faults[k][2], " ($k) > ", length(contiguous[k]), " Events, MTTR: ", round(mean(contiguous[k]), 1), ", Event Lengths: ", contiguous[k])
            println("MTBF: ", avail / (sum(contiguous[k])+1))
        else
            #println(faults[k][1], " - ", faults[k][2], " ($k) > 0 Events")
        end
    end
end

#importDailies(DateTime(2019, 5, 20))
#availabilityColour(DateTime(2019, 5, 20), now())
println("Availability Events")
for line in ["Auto2"] # ["Auto1", "Auto2", "EB", "Flexi", "HV", "Paint"]
    MTBF(line, DateTime(2010,5,20,0,0,0), DateTime(2019, 5, 30, 0,0,0))
    println()
end
