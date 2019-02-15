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
    exebind!("PABi", vals)
    last_insert(PabDB)
end

insertAvailLoss!(vals) = exebind!(vals)

line(handle) = try SQLite.query(PabDB, "SELECT Line from Lines WHERE Handle=?", values=[handle])[1][1] catch "" end

function faultID(line, stage, fault)
    try
        return SQLite.query(PabDB, "SELECT Fault_ID FROM Faults WHERE Line=? AND Stage=? AND Fault=?", values=[line, stage, fault])[1][1]
    end
    exebind!("Faulti", [line, stage, fault])
    SQLite.query(PabDB, "SELECT last_insert_rowId()")[1][1]
end
###############
end
