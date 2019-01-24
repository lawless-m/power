module PAB

using ExcelReaders

export readweeks

function readweeks(io, line, dir)
    println(dir)
    for wk in filter((f)->f[1] == 'W' && f[end-3:end] == ".xls", readdir(dir))
        wknum = wk[6:end-4]
        sht = sheet(dir, wk, "Quality")
        c = 3
        while typeof(sht[2,c]) == String
            println(io, wknum, "\t", sht[2,c], "\t", sht[3,c])
            c += 1
        end
    end
end

sheet(xld, xlfn, tab) =  readxlsheet(xld * "\\" * xlfn, tab)

end
