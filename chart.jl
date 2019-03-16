
d = Dict("Auto 1"=>Dict("W7"=>(2,7), "W8"=>(3,1), "W9"=>(3,3), "W10"=>(5,3)),
"EB"=>Dict("W7"=>(7,2), "W8"=>(5,0), "W9"=>(7,1), "W10"=>(2,3)),
"Auto 2"=>Dict("W7"=>(9,0), "W8"=>(0,0), "W9"=>(5,0), "W10"=>(0,0)),
"Flex"=>Dict("W7"=>(0,4), "W8"=>(5,0), "W9"=>(3,1), "W10"=>(1,1)),
"Coating"=>Dict("W7"=>(3,2), "W8"=>(4,0), "W9"=>(4,0), "W10"=>(1,0)),
"HV"=>Dict("W7"=>(4,0), "W8"=>(1,0), "W9"=>(3,0), "W10"=>(3,0)))

println("[")
for line in ["Auto 1", "EB", "Auto 2", "Flex", "Coating", "HV"]
    v = d[line]
    for wk in 7:10
        t = v["W$wk"]
        println("[\"$line\", \"W$wk\", \"done\", $(t[1]), $(t[2])],")
        println("[\"$line\", \"W$wk\", \"undone\", $(t[1]), $(t[2])],")
    end
    if line == "HV"
        println("]")
    else
        println("[],")
    end
end
