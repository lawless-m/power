
include("../../credentials.jl")
include("dirs.jl")

using Plexus
using PMDB

plex = Plex(credentials["plex"]...)

fill_equipment(equipment(plex, ))

