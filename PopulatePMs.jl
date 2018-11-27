
include("../../credentials.jl")
include("dirs.jl")

using PMDB
using Cache
using Plexus

plex = Plex(credentials["plex"]...)

fill_equipment(equipment(plex, lines()))

clear_pm_tasks()
clear_pms()

for pm in cache("pm_list", ()->pm_list(plex))
	insert_pm(parse(Int, pm.checklist_no)
			, parse(Int, pm.checklist_key)
			, parse(Int, pm.equipkey)
			, pm.pmtitle
			, pm.maintfrm.priority
			, parse(Int, pm.freq)
			, replace(pm.maintfrm.sinstruct, ['\r', '\n'], "")
			, pm.maintfrm.hours
			, Dates.value(pm.start)
			, pm.maintfrm.tasklist
			)
end

