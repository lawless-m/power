include("dirs.jl")

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

function homes()

  for line in ["EB 2018", "Line 5 2018", "HV 2018", "Autoline 2018"]
    io = open("$home\\Weekly PAB\\$line.txt", "w+")
    readweeks(io, line, "$home\\Weekly PAB\\$line")
    close(io)
  end
end

function aways()
  io = open("N:\\Weekly PAB-OEE Data\\Auto Line 2 Volvo.txt", "w+")
  readweeks(io, "Auto Line2 2018","N:\\Weekly PAB-OEE Data\\Auto Line 2 Volvo\\Auto Line2 2018")
end

aways()
