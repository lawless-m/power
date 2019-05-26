
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

function process_faults(earliest, latest, events)

    contiguous = Vector{Int}()
    tbf = Vector{Int}()

    loss = 0
    started = ended = earliest

    for d in events # d = (startt, endt, avail, loss)

        println("$d - $loss")

        if d[3] == d[4] # whole of this period was lost
            println("p1 $loss")
            ended = min(ended, d[1])
            if ended == d[1] # this event ended the timer
                println("p1a $loss")
                push!(tbf, Dates.value(Dates.Minute(ended - started)))
            else

                println("p1b $loss")
                # this event was over two contiguous periods and already ended
            end
            started = d[2]
            loss += d[4]
            continue
        end

        # last event ended when this one started so put it at the start of the period, add this loss on to the end then restart timer
        if d[1] == started
            println("p2 $loss")
            loss += d[4]
            push!(contiguous, loss)
            loss = 0
            started += Dates.Minute(d[4])
            continue
        end

        if loss > 0 # previous loss was whole period and not contiguous with this one
            println("p3 $loss")
            started = d[1]
            ended = d[2]
            push!(contiguous, loss)
            loss = 0
            continue
        end

        println("p4 $loss")
        # this is a new event, put it at the end of the period
        ended = d[2] - Dates.Minute(d[4])
        push!(tbf, Dates.value(Dates.Minute(ended - started)))
        started = d[2]
        ended = d[2]
        loss = d[4]

    end

    if ended > started
        push!(tbf, Dates.value(Dates.Minute(ended - started)))
    end
    push!(tbf, Dates.value(Dates.Minute(latest - started)))
    if loss > 0 # previous event was the last
        println("p5 $loss")
        push!(contiguous, loss)
    end

    contiguous, tbf
end

function print_results(faults, contiguous, tbf)
    for k in keys(faults)
        avail = sum(tbf[k])
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

function event_list(faultids, slots)
    events = Dict{Int, Vector{Tuple{DateTime, DateTime, Int, Int}}}()
    foreach(k->events[k] = Vector{Int}(), faultids)
    for r in 1:size(slots, 1)
        if typeof(slots[r, :loss]) != Missings.Missing
            push!(events[slots[r,:fault_id]], (int2time(slots[r, :startT]), int2time(slots[r, :endT]), 60 - slots[r, :stopmins], slots[r, :loss]))
        end
    end
    events
end

function group_slots(slots)
    grouped = Dict{DateTime, Tuple{DateTime, Int, Vector{Tuple{Int, Int}}}}()
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
                grouped[st] = (slot_st, actual, [])
            end
            endt, actual, events = get(grouped, slot_st, (slot_et, slots[r, :actual], []))
            push!(events, (slots[r, :fault_id], slots[r, :loss]))
            grouped[slot_st] = (endt, actual, events)
            actual = 0
            st = slot_et
        end
    end
    if actual > 0
        grouped[st] = (et, actual, [])
    end
    grouped
end

function MTBF(line, st, se)
    slots = PABDB.all_slots(line, st, se) #  line, startT, endT, stopmins, loss, fault_id
    if size(slots, 1) == 0
        return
    end
    group_slots(slots)
    exit(0)

    faults = PABDB.idfaults(line)

    earliest = int2time(slots[1, :startT])
    latest = int2time(slots[end, :startT])

    events = event_list(keys(faults), slots)
    contiguous = Dict{Int, Vector{Int}}()
    tbf = Dict{Int, Vector{Int}}()
    for f in keys(faults)
        println("F $f")
        contiguous[f], tbf[f] = process_faults(earliest, latest, events[f])
    end

    println(line, "\n From $st to $se")
    print_results(faults, contiguous, tbf)
end

#importDailies(DateTime(2019, 5, 20))
#availabilityColour(DateTime(2019, 5, 20), now())
println("Availability Events")
for line in ["Auto1", "Auto2", "EB", "Flexi", "HV", "Paint"]
    MTBF(line, DateTime(2010,5,20,0,0,0), DateTime(2019, 5, 30, 0,0,0))
    println()
end
