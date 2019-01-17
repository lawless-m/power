
include("dirs.jl")

struct Node
	name::String
	parent::Union{Node, Void}
	children::Vector{Node}
	function Node(p, n) 
		new(n, p, Vector{Node}())
	end
end

function child!(p, s)
	nd = Node(p, s)
	push!(p.children, nd)
	nd
end

function buildTree(io)
	root = Node(nothing, "AAM")
	while ((l=readline(io)) != "")
		println(l)
		bits = split(l, "\t")
		inds = length(bits)
		if inds == 1
			build = child!(root, bits[end])
			continue
		end
		if inds == 2
			line = child!(build, bits[end])
			continue
		end
		if inds == 3
			mach = child!(line, bits[end])
			continue
		end
			
		eqs = split(bits[end], " - ")
		if length(eqs) == 1
			submach = child!(mach, bits[end])
			continue
		end
	end
	root
end


function writeTree(aam)
	grph = raw"Z:\Maintenance\Matt-H\EquipTree\tabbed.dot"
	dir = raw"Z:\Maintenance\Matt-H\EquipTree"
	#println(io, "a0 [label=\"", aam.name, "\"; shape=\"rarrow\"]")
	a = b = c = d = e = f = 0
	for build in aam.children
	#	println(io, "b0 [label=\"", build.name, "\"; shape=\"house\"]")
	#	println(io, "a$a->b$b")
		for line in build.children
			io = open("$dir\\$(line.name).dot", "w+")
			println(io, "digraph Equip {")
			println(io, "rankdir=LR;")
			println(io, "c$c [label=\"", line.name, "\"; shape=\"cds\"]")
	#		println(io, "b$b->c$c")
			for equip in line.children
				println(io, "d$d [label=\"", equip.name, "\"; shape=\"octagon\"]")
				println(io, "c$c->d$d")
				for ass in equip.children
					println(io, "e$e [label=\"", ass.name, "\"; shape=\"box\"]")
					println(io, "d$d->e$e")
					for subass in ass.children
						println(io, "f$f [label=\"", subass.name, "\"; 	shape=\"box\"]")
						println(io, "e$e->f$f")
					end
					e += 1
				end
				d += 1
			end
			println(io, "}")
			close(io)
			c += 1
		end
		b += 1
	end
end

tree = open(buildTree, raw"Z:\Maintenance\Matt-H\EquipTree\tabbed.txt", "r")
writeTree(tree)
