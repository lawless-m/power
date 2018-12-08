module EBDB

using SQLite
using SQLiteTools

export cleardb, insertEB, allEB, file_recorded, last_inum

include("dirs.jl")

EBdb = SQLite.DB("$dbdir\\EB_Perf.db")

println("$dbdir\\EB_Perf.db")

cleardb() = truncate!(EBdb, "EB")

SQLiteTools.insert!(EBdb, "EBi", "Insert into EB (Date, Leader, Shift, Line, Product, Std_Rate_PPH, Avail_Hours, Product_Max, Product_Avail, Product_Variance, Time_Variance, Efficiency, Item, Process, Problem, Loss_Mins, Effect, OEE_Element, Action, Fix_Or_Repair, Weld_section_due, IncidentNo, filename) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

insertEB(vals) = exebind!("EBi", vals)
allEB() =  SQLite.query(EBdb, "select * from EB")

file_recorded(fn) = SQLite.query(EBdb, "select count(*) from EB WHERE filename=?", values=[fn])[1][1]>0

last_inum() = SQLite.query(EBdb, "select max(IncidentNo) from EB")[1][1]

end
