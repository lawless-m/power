
include("dirs.jl")

using SQLiteTools
using XlsxWriter
using PAB

struct Wb
    wb
    fmts::Dict
    function Wb()
        wb = Workbook(joinpath(Home, "OEE", "Avail.xlsx"))
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
    slot_times = sort(collect(keys(slots)))
    for id in faults
        ud[id] = Vector{Union{uptime, downtime}}()
        push!(ud[id], uptime(slot_times[1], slot_times[1], 0))
    end
    for slot_st in slot_times
        slot_et, actual, events = slots[slot_st]
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

function print_grouped(grouped)

end

function MTBF(line, st, se)
    slots = PABDB.all_slots(line, st, se) #  line, startT, endT, stopmins, loss, fault_id
    if size(slots, 1) == 0
        return
    end
    grouped = group_slots(slots)
    # foreach(st->println(st, grouped[st]), sort(collect(keys(grouped)))

    faults = PABDB.idfaults(line)
    ud = updowns(grouped, keys(faults))

    setrounding(Float64, Base.RoundUp)
    for id in keys(ud)
        print(faults[id])
        #println(ud[id])
        ut = Dates.value(sum(duration, ud[id][isup.(ud[id])]))
        print(" Uptime: ", ut, " mins")
        if length(ud[id]) > 1
            dt = Dates.value(sum(duration, ud[id][isdown.(ud[id])]))
            cnt = count(isdown.(ud[id]))
            print(" Downtime: ", dt, " mins")
            print(" Events: ", cnt)
            @printf(" MTBF: %0.1f mins, MTTR: %0.1f mins", round(ut / cnt, 1), round(dt / cnt, 1))
        end
        println()
    end
end

#importDailies(DateTime(2019, 5, 20))
#availabilityColour(DateTime(2019, 5, 20), now())
println("Availability Events")
months = Dict()
for line in ["Auto1", "Auto2", "EB", "Flexi", "HV", "Paint"]
    MTBF(line, DateTime(2010,1,1,0,0,0), DateTime(2019, 2, 1, 0,0,0))
    println()
end
