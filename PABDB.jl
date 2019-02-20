module PABDB

include("dirs.jl")

using SQLite
using SQLiteTools

export insertPAB!, faultID, insertAvailLoss!, insertPAB!

PabDB = SQLite.DB("$dbdir\\PABDB.db")

SQLiteTools.insert!(PabDB, "PABi", "INSERT INTO PAB (Line, StartT, EndT, Reason, StopMins, Part, Target, Operator, Actual, Comment) VALUES(?,?,?,?,?,?,?,?,?,?)")
SQLiteTools.insert!(PabDB, "Faulti", "INSERT INTO Faults (Line, Stage, Fault) VALUES(?,?,?)")
SQLiteTools.insert!(PabDB, "AvailLossi", "INSERT INTO AvailabilityLoss (PAB_ID, Fault_ID, Loss) VALUES(?,?,?)")

function insertPAB!(vals)
    if size(SQLite.query(PabDB, "SELECT Line FROM PAB WHERE Line=? AND StartT=? AND EndT=?", values=[vals[1], vals[2], vals[3]]), 1) > 0
        return 0
    end
    exebind!("PABi", vals)
    return last_insert(PabDB)
end

insertAvailLoss!(vals) = exebind!("AvailLossi", vals)

line(handle) = try SQLite.query(PabDB, "SELECT Line from Lines WHERE Handle=?", values=[handle])[1][1] catch "" end

function faultID(line, stage, fault)
    try
        return SQLite.query(PabDB, "SELECT Fault_ID FROM Faults WHERE Line=? AND Stage=? AND Fault=?", values=[line, stage, fault])[1][1]
    end
    exebind!("Faulti", [line, stage, fault])
    SQLite.query(PabDB, "SELECT last_insert_rowId()")[1][1]
end

function faultList()
    faults = Dict{String, Dict{String, Dict{String, Int}}}()
    flts = SQLite.query(PabDB, "SELECT * FROM Faults")
    for r in 1:size(flts, 1)
        ld = get!(faults, flts[r, :Line], Dict{String, Dict{String, Int}}())
        sd = get!(ld, flts[r, :Stage], Dict{String, Int}())
        sd[flts[r, :Fault]] = flts[r, :Fault_ID]
    end
    faults
end

function pabsBetween(s::DateTime, e::DateTime)
    SQLite.query(PabDB, "SELECT * FROM PAB WHERE StartT >=? AND EndT <=? order by Line, StartT desc", values=[Dates.value(s), Dates.value(e)], stricttypes=true)
end

function availLossBetween(s::DateTime, e::DateTime)
    SQLite.query(PabDB, "SELECT AvailabilityLoss.* FROM PAB JOIN AvailabilityLoss on PAB.id=AvailabilityLoss.PAB_ID  WHERE StartT >=? AND EndT <=?", values=[Dates.value(s), Dates.value(e)])
end

###############
end
