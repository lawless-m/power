
include("dirs.jl")

using SQLiteTools
using XlsxWriter
using PAB
using Weibull

struct Wb
    wb
    fmts::Dict
    function Wb(fn)
        wb = Workbook(fn)
        fmts = Dict()
        fmts["angle"] = add_format!(wb, Dict("rotation"=>45))
        fmts["day_fmt"] = add_format!(wb, Dict("num_format"=>"dd/mm/yy"))
        fmts["time_fmt"] = add_format!(wb, Dict("num_format"=>"hh:mm"))
        fmts["round"] = add_format!(wb, Dict("num_format"=>"0"))
        new(wb, fmts)
    end
end

function availabilityColour(startt, endt)
    println("Writing to ", joinpath(Home, "OEE", "Avail.xlsx"))
    wb = Wb(joinpath(Home, "OEE", "Avail.xlsx"))
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


    if ! (line in keys(flts))
        return
    end

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


struct uptime
    up::DateTime
    down::DateTime
    actual::Int
end

struct downtime
    down::DateTime
    up::DateTime
end

isup(x) = typeof(x) == uptime
isdown(x) = typeof(x) == downtime
duration(u::uptime) = Dates.Minute(u.down - u.up)
duration(d::downtime) = Dates.Minute(d.up - d.down)

function group_slots(slots)
    grouped = Dict{DateTime, Tuple{DateTime, Int, Dict{Int, Tuple{Int, Int}}}}()
    # st=>et, Actual, [(faultID, loss)...]

    st = int2time(slots[1, :startT]) # for Min below
    et = st
    actual = 0
    for r in 1:size(slots, 1)
        slot_st = int2time(slots[r, :startT])
        slot_et = int2time(slots[r, :endT])
        if slot_et < slot_st
            slot_et += Dates.Day(1)
        end
        if typeof(slots[r, :loss]) == Missings.Missing
            st = min(st, slot_st)
            et = slot_et
            actual += slots[r, :actual]
        else
            if st != slot_st
                grouped[st] = (slot_st, actual, Dict{Int, Int}())
            end
            endt, actual, events = get(grouped, slot_st, (slot_et, slots[r, :actual], Dict{Int, Tuple{Int, Int}}()))
            events[slots[r, :fault_id]] = (slots[r, :stopmins], slots[r, :loss])
            grouped[slot_st] = (endt, actual, events)
            actual = 0
            st = slot_et
        end
    end
    if actual > 0
        grouped[st] = (et, actual, Dict{Int, Tuple{Int, Int}}())
    end
    grouped
end

function updowns(slots, faults)
    ud = Dict{Int, Vector{Union{uptime, downtime}}}()
    gslots = group_slots(slots)
    slot_times = sort(collect(keys(gslots)))
    for id in faults
        ud[id] = Vector{Union{uptime, downtime}}()
        push!(ud[id], uptime(slot_times[1], slot_times[1], 0))
    end
    for slot_st in slot_times
        slot_et, actual, events = gslots[slot_st]
        #println(slot_st, " - ", slot_et)
        for id in faults
            if id in keys(events)
                if typeof(ud[id][end]) == uptime # went down
                    #println(id, " - went down")
                    dt = slot_et - Dates.Minute(events[id][2])
                    ud[id][end] = uptime(ud[id][end].up, dt, ud[id][end].actual + actual)
                    push!(ud[id], downtime(dt, slot_et))
                else # stayed down
                    #println(id, " - stayed down")
                    ut = slot_st + Dates.Minute(events[id][1] + events[id][2])
                    ud[id][end] = downtime(ud[id][end].down, ut)
                    if ut < slot_et
                        push!(ud[id], uptime(ut, slot_et, actual))
                    end
                end
            else
                if typeof(ud[id][end]) == uptime # stayed up
                    #println(id, " - stayed up")
                    ud[id][end] = uptime(ud[id][end].up, slot_et, ud[id][end].actual + actual)
                else # went up
                    #println(id, " - went up")
                    ud[id][end] = downtime(ud[id][end].down, slot_st)
                    push!(ud[id], uptime(slot_st, slot_et, actual))
                end
            end
        end
    end
    ud
end

struct Stat
    start::DateTime
    finish::DateTime
    parts::Int
    events::Int
    uptime::Int
    downtime::Int
end


function stats_by_week_workbook(faults, stats, path)
    wb = Wb(path)

    for line in sort(collect(keys(stats)))
        faultname(id) = faults[line][id][1] == "Equipment" ? faults[line][id][2] : "$(faults[line][id][1]) - $(faults[line][id][2])"
        ws = add_worksheet!(wb.wb, line)

        function faultsort(a, b)
            if faults[line][a][1] == faults[line][b][1]
                return faults[line][a][2] < faults[line][b][2]
            end
            faults[line][a][1] < faults[line][b][1]
        end
        coff = minimum(faults[line])[1]-2
        numr = length(collect(keys(stats[line])))
        r = 2
        write!(ws, r, 0, "MTBF")
        write!(ws, r+numr+1, 0, "MPBF")
        write!(ws, r+2numr+2, 0, "MTTR")
        write!(ws, r+3numr+3, 0, "Events")
        fst = true
        for date in sort(collect(keys(stats[line])))

            W = "W$(Dates.week(date))"
            write!(ws, r, 1, W, wb.fmts["day_fmt"])
            write!(ws, r+numr+1, 1, W, wb.fmts["day_fmt"])
            write!(ws, r+2numr+2, 1, W, wb.fmts["day_fmt"])
            write!(ws, r+3numr+3, 1, W, wb.fmts["day_fmt"])
            for id in keys(stats[line][date])
                stat = stats[line][date][id]
                if fst
                    write!(ws, 0, id-coff, faults[line][id][1])
                    write!(ws, 1, id-coff, faults[line][id][2])
                end
                if stat.downtime > 0
                    write!(ws, r, id-coff, stat.uptime / stat.events, wb.fmts["round"])
                    write!(ws, r+numr+1, id-coff, stat.parts / stat.events, wb.fmts["round"])
                    write!(ws, r+2numr+2, id-coff, stat.downtime / stat.events, wb.fmts["round"])
                    write!(ws, r+3numr+3, id-coff, stat.events, wb.fmts["round"])
                end
            end
            fst = false
            r += 1
        end
    end
    close(wb.wb)
end

function statsummary(faults, ud, st, se)
    stats = Dict{Int, Stat}()

    for id in keys(faults)
        ut = Dates.value(sum(duration, ud[id][isup.(ud[id])]))
        parts = sum(e->e.actual, ud[id][isup.(ud[id])])
        if length(ud[id]) > 1
            dt = Dates.value(sum(duration, ud[id][isdown.(ud[id])]))
            cnt = count(isdown.(ud[id]))
            stats[id] = Stat(st, se, parts, cnt, ut, dt)
        else
            stats[id] = Stat(st, se, parts, 0,0,0)
        end
    end
    stats
end

function stats_by_days(faultsbyline, wkst, wkend, days=7)
    stats = Dict{String, Dict{DateTime, Dict{Int, Stat}}}()
    for line in ["Auto1", "Auto2", "EB", "Flexi", "Paint", "HV"]
        stats[line] = Dict{DateTime, Dict{Int, Stat}}()
        for se in wkst+Dates.Day(days):Dates.Day(7):wkend
            st = se - Dates.Day(days)
            slots = PABDB.all_slots(line, st, se) #  line, startT, endT, stopmins, loss, fault_id
            if size(slots, 1) > 0
                stats[line][se - Dates.Day(1)] = statsummary(faultsbyline[line], updowns(slots, keys(faultsbyline[line])), st, se)
            end
        end
    end
    stats
end


function stats_to_week(faultsbyline, wkst, wkend)
    stats = Dict{String, Dict{DateTime, Dict{Int, Stat}}}()
    for line in ["Auto1", "Auto2", "EB", "Flexi", "Paint", "HV"]
        stats[line] = Dict{DateTime, Dict{Int, Stat}}()
        for st in wkst:Dates.Day(7):wkend
            se = st + Dates.Day(7)
            slots = PABDB.all_slots(line, wkst, se) #  line, startT, endT, stopmins, loss, fault_id
            if size(slots, 1) > 0
                stats[line][st] = statsummary(faultsbyline[line], updowns(slots, keys(faultsbyline[line])), st, se)
            end
        end
    end
    stats
end


function survivors(faultsbyline, wkst, wkend)
    weibulls = Dict{String, Dict{Int, Tuple{Float64, Float64}}}()
    for line in ["Auto1", "EB", "Flexi", "Paint", "HV"] #"Auto2"
        faultids = keys(faultsbyline[line])
        weibulls[line] = Dict{Int, Tuple{Float64, Float64}}()
        slots = PABDB.all_slots(line, wkst, wkend) #  line, startT, endT, stopmins, loss, fault_id
        if size(slots, 1) > 0
            ud = updowns(slots, faultids)
            for id in faultids
                uptimes = Vector{Float64}()
                for u in ud[id][isup.(ud[id])]
                    t = Dates.value(Dates.Minute(u.down-u.up))
                    if t > 0
                        push!(uptimes, t)
                    end
                end
                if length(uptimes) > 1
                    try
                        weibulls[line][id] = Weibull.fit(uptimes)
                        println(line, "\t", id, "\t", faultsbyline[line][id][1], "\t", faultsbyline[line][id][2], "\t", weibulls[line][id][1], "\t", weibulls[line][id][2])
                    catch
                        #println("ERROR ", uptimes)
                        weibulls[line][id] = (0,0)
                        println(line, "\t", id, "\t", faultsbyline[line][id][1], "\t", faultsbyline[line][id][2], "\t0\t0")
                    end
                else
                    weibulls[line][id] = (0,0)
                    println(line, "\t", id, "\t", faultsbyline[line][id][1], "\t", faultsbyline[line][id][2], "\t0\t0")
                end
            end
        end
    end
    exit()
end

#PAB.importDailies(DateTime(2019, 6, 19)) # since
#availabilityColour(DateTime(2019, 6, 1), now())


lts = PAB.PABDB.latest_PABs()
for r in 1:size(lts, 1)
    println(lts[r, 1], ": ", Date(int2time(lts[r,2])))
end

#wkst = DateTime(2018,12,31,0,0,0)

wkst = DateTime(2019,1,1,0,0,0)
wkend = now()
faultsbyline = PABDB.idfaults()

survivors(faultsbyline, wkst, wkend)


stats = stats_by_days(faultsbyline, wkst, wkend, 28)
stats_by_week_workbook(faultsbyline, stats, joinpath(Home, "MTBF_28day.xlsx"))

stats = stats_by_days(faultsbyline, wkst, wkend, 7)
stats_by_week_workbook(faultsbyline, stats, joinpath(Home, "MTBF_7day.xlsx"))

stats = stats_to_week(faultsbyline, wkst, wkend)
stats_by_week_workbook(faultsbyline, stats, joinpath(Home, "MTBF_cumulative.xlsx"))
