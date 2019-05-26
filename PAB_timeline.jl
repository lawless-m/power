
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
