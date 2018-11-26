module EBDB

using SQLite
using SQLiteTools

export cleardb, insertEB, allEB

if isfile(raw"c:\users\matt\ntuser.ini")
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
else
	dbdir = raw"Z:\Maintenance\Matt-H\power\SQLite.Data"
end

EBdb = SQLite.DB("$dbdir\\EB_Perf.db")

println("$dbdir\\EB_Perf.db")

cleardb() = truncate!(EBdb, "EB")

SQLiteTools.insert!(EBdb, "EBi", "Insert into EB (Date, Leader, Shift, Line, Product, Std_Rate_PPH, Avail_Hours, Product_Max, Product_Avail, Product_Variance, Time_Variance, Efficiency, Item, Process, Problem, Loss_Mins, Effect, OEE_Element, Action, Fix_Or_Repair, Weld_section_due, IncidentNo) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

insertEB(vals) = exebind!("EBi", vals)
allEB() =  SQLite.query(EBdb, "select * from EB")

end