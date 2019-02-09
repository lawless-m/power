include("dirs.jl")

using PAB


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
