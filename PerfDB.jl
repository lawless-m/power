module PerfDB

using SQLite
using SQLiteTools

export clearEB, insertEB, allEB, file_recorded, last_inum

include("dirs.jl")

perfDB = SQLite.DB("$dbdir\\Perf.db")

println("$dbdir\\Perf.db")

clear(line) = truncate!(perfDB, "EB")

SQLiteTools.insert!(perfDB, "EBi", "Insert into EB (Date, Leader, Shift, Line, Product, Std_Rate_PPH, Avail_Hours, Product_Max, Product_Actual, Product_Variance, Time_Variance, Efficiency, Item, Process, Problem, Loss_Mins, Effect, OEE_Element, Action, Fix_Or_Repair, Weld_section_due, IncidentNo, filename) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

SQLiteTools.insert!(perfDB, "PaintI", "Insert into Paint (Date, Leader, Shift, Product, Std_Rate_PPH, Avail_Hours, Product_Max, Product_Actual, Product_Variance, Time_Variance, Efficiency, Line, Item, Process, Problem, Quality_Defect_Type, Quality_Lost_Parts, Loss_Mins, Effect, OEE_Element, Action, Fix_Or_Repair, IncidentNo, filename) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

insertEB(vals) = exebind!("EBi", vals)
allEB() =  SQLite.query(perfDB, "select * from EB")

insertPaint(vals) = exebind!("PaintI", vals)

file_recorded(fn, line) = SQLite.query(perfDB, "select count(*) from $line WHERE filename=?", values=[fn])[1][1]>0

last_inum(line) = SQLite.query(perfDB, "select max(IncidentNo) from $line")[1][1]

end
