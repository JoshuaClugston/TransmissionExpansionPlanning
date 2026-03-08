# using Pkg; Pkg.add("XLSX"); Pkg.add("CPLEX"); Pkg.add("JuMP"); Pkg.add("HiGHS") # incase not installed already 
using JuMP, CPLEX, HiGHS, XLSX

data = XLSX.readxlsx("./data/data.xlsx") # xlsx with data in separate sheets 

bus = data["bus"]
branch = data["branch"] # set of existing lines 
gen = data["gen"]
new_branch = data["new_branch"] 

# bus data
bus_i = [] # bus id
d = [] # demand at bus 

# line data
line_existing_i = [] # existing line index
fbus_existing = [] # fbus existing
tbus_existing = [] # tbus existing
F_existing = [] # line thermal limit -- existing
B_existing = [] # susceptance -- existing 

# gen data
gen_i = [] # generator index 
gen_lim = [] # max production level
gen_cost = []
gen_bus = []

# assign values to bus data 
for row in XLSX.eachrow(bus)
    rn = XLSX.row_number(row) # row number
    if rn == 1
        continue
    else
        push!(bus_i, row[1]) # index for buses (bus ID)
        push!(d, row[11])
        # push!(d, row[4])
    end
end

# assign values to line/branch data
for (i,row) in enumerate(XLSX.eachrow(branch)) # existing lines
    rn = XLSX.row_number(row) # row number
    if rn == 1
        continue
    else
        push!(line_existing_i, i-1) 
        push!(fbus_existing, row[2])
        push!(tbus_existing, row[3])
        push!(B_existing, 1/row[9]) # susceptance
        # push!(B_existing, row[10]) # susceptance 
        push!(F_existing, row[11]) ## using continuous rating 
    end
end

function get_new_line()
        count = length(line_existing_i) # set the initial count to the length of the existing lines -- should be 120
        line_new_i = []
        fbus_new = []
        tbus_new = []
        cost_new = []
        B_new = []
        F_new = []

        for (i,row) in enumerate(XLSX.eachrow(new_branch)) # new lines 
            rn = XLSX.row_number(row) # row number
            if rn == 1
                continue
            else
                if row[19] != -1 # ignore lines which have no line installed
                    count += 1
                    push!(line_new_i, count) # index for new lines -- starting at the end o the existing lines
                    push!(fbus_new, row[2]) 
                    push!(tbus_new, row[3])
                    push!(cost_new, row[19]) ## cost for installing new line -- 900k per mile 
                    # push!(B_new, row[10]) 
                    push!(B_new, 1/row[9]) 
                    # push!(F_new, row[17]) 
                    push!(F_new, row[11])
                end
            end
        end

        return line_new_i, fbus_new, tbus_new, cost_new, B_new, F_new
end

line_new_i, fbus_new, tbus_new, cost_new, B_new, F_new = get_new_line()

#assign values to gen data
function get_generators()
    counter = 0
    for (i,row) in enumerate(XLSX.eachrow(gen))
        rn = XLSX.row_number(row) # row number
        if rn == 1
            continue
        else
            if row[10] != -1
                counter += 1 
                push!(gen_i, counter) # index for generators 
                push!(gen_lim, row[10])
                push!(gen_cost, row[9])
                push!(gen_bus, row[1])
                # println(row[2])
            else
                counter += 1 
                push!(gen_i, counter) # index for generators 
                push!(gen_lim, row[10]*0)
                push!(gen_cost, row[9])
                push!(gen_bus, row[1])
            end
        end 
    end
    return gen_i, gen_lim, gen_cost
end


gen_i, gen_lim, gen_cost = get_generators()
println("length of ", length(gen_i))


# model = Model(HiGHS.Optimizer) # empty model using HiGHS
model = Model(CPLEX.Optimizer) # empty model using HiGHS

d = 3*d # multiply the load/demand at each bus by 3

### define parameters 

M = 100000.0 # large M


### define variables 

w = Dict(); # dictionary for new line installation decision variables
P_gen = Dict(); # dictionary for generation variables
δ = Dict(); # dictionary for phase angle variables
P_flow = Dict(); # dictionary for power flow variables


for (k,i,j) in zip(line_existing_i, fbus_existing, tbus_existing) 
    P_flow[k,i,j] = @variable(model, lower_bound = -F_existing[k], upper_bound = F_existing[k], base_name = "P_flow[$(k),$(i), $(j)]") ### existing lines 
end

for ((idx,k),i,j) in zip(enumerate(line_new_i), fbus_new,tbus_new) 
    if cost_new[idx] != 0
        P_flow[k,i,j] = @variable(model, lower_bound = -F_new[idx], upper_bound = F_new[idx], base_name = "P_flow[$(k),$(i),$(j)]") ### new lines
    else
        println("removing bounds for transformers")
        P_flow[k,i,j] = @variable(model, base_name = "P_flow[$(k),$(i),$(j)]") ### new lines
    end
end


for k in line_new_i
    w[k] = @variable(model, base_name="w[$(k)]", binary=true)
end

for n in bus_i 
    δ[n] = @variable(model, base_name="δ[$(n)]")
end 


for (g,n) in zip(gen_i,gen_bus)
    P_gen[g,n] = @variable(model, lower_bound = 0, upper_bound = gen_lim[g], base_name= "P_gen[$(g), $(n)]") 
end

println("amount of generation variables ", length(P_gen))

println("number of variables = ", length(P_gen) + length(δ) + length(w) + length(P_flow))
println("number of binary = ", length(w))
 
##### define constraints 
for (g,n) in zip(gen_i,gen_bus)  
    @constraint(model, P_gen[g,n] <= gen_lim[g])
end

for (k,s,r) in zip(line_existing_i, fbus_existing, tbus_existing)
    @constraint(model, P_flow[k,s,r] == B_existing[k]*(δ[r] - δ[s])) ###????
end

for ((i,k),s,r) in zip(enumerate(line_new_i), fbus_new, tbus_new)
    if cost_new[i] != 0 # ignore bounds for transformers since we assume infinite bounds 
        @constraint(model, P_flow[k,s,r] >= -F_new[i]*w[k]) 
        @constraint(model, P_flow[k,s,r] <= F_new[i]*w[k])  
    end
    @constraint(model, P_flow[k,s,r] - B_new[i]*(δ[r]-δ[s]) >= -(1-w[k])*M) 
    @constraint(model, P_flow[k,s,r] - B_new[i]*(δ[r]-δ[s]) <= (1-w[k])*M) 
    # end
end

constraint = Dict();

for (i,n) in enumerate(bus_i) 
    SUM_LHS = 0

    for (k, fbus, tbus) in zip(line_existing_i, fbus_existing, tbus_existing)
        if fbus == n
            SUM_LHS += P_flow[k,n,tbus] ## get the second term in the left hand side: k = n^r (receiving bus)
        end
        if tbus == n
            SUM_LHS -= P_flow[k,fbus,n] ## get the first term in the left hand side: k = n^s (sending bus)
        end
    end


    for (k,s,r) in zip(line_new_i, fbus_new, tbus_new)
        if s == n
            SUM_LHS += P_flow[k,n,r] ## get the second term in the left hand side: k = n^r (receiving bus)
        end
        if r == n
            SUM_LHS -= P_flow[k,s,n] ## get the first term in the left hand side: k = n^s (sending bus)
        end
    end

    SUM_GEN = 0 # for right hand side 
    for (g,val) in zip(gen_i,gen_bus)
        if val == n
            SUM_GEN += P_gen[g,n] # get the right hand side value 
        end 
    end

    constraint[i] = @constraint(model, SUM_LHS == SUM_GEN - d[i])
end

# println(sort(collect(constraint), by = x->x[1])) # print balance constraint sorted in order of keys

# ### define objective function 

@objective(model, Min, sum(gen_cost[g]*P_gen[g,n] for (g,n) in zip(gen_i,gen_bus)) + sum( cost_new[i]*w[k] for (i,k) in enumerate(line_new_i)))

println("number of constraints ", sum(num_constraints(model, F, S) for (F, S) in list_of_constraint_types(model)))

optimize!(model)
println(solution_summary(model, verbose = true))
# println("objective value = ", objective_value(model))

new_line_installed = []
for k in line_new_i
    if value(w[k]) > .75
        println(value(w[k]))
        println("index in new lines: ", k-120)
        if k-120 < 120
            println("index original, without modification: ", k)
            push!(new_line_installed,k-120)
        end
    end
end

println(new_line_installed)
println(length(new_line_installed)) 


let costs_summation = 0
    for k in new_line_installed
        println(k, " ", cost_new[k])
        costs_summation = costs_summation + cost_new[k]
    end
    println("total costs:", costs_summation)
end



gen_running = []
for (g, n, l) in zip(gen_i, gen_bus, gen_lim)
    if g > 99
        push!(gen_running, (g,value(P_gen[g,n]), l))
    end
end
# gen_running = [(g,value(P_gen[g,n])) for (g,n) in zip(gen_i, bus_i)]

let count1 = 0, count2 = 0, count3 = 0, count4 = 0, count5 = 0, count6 = 0
    for i in gen_running 
        println("(generator id, variable value, limit) = ", i)
        if trunc(Int,i[2]) == i[3]
            count1 += 1
        end
        if i[2] == 0.0 && trunc(Int,i[2]) != i[3]
            count2 += 1
        end
        if i[2] < i[3] && i[2] != 0.0
            count3 += 1
        end
    end
    println("total number of new generators using their total capacity ", count1)
    println("total number of new generators not using any of their total capacity ", count2)
    println("total number of new generators not using all of their total capacity ", count3)

    for (g,n,l) in zip(gen_i, gen_bus, gen_lim)
        if trunc(Int,value(P_gen[g,n])) == l
            count4 += 1
        end
        if value(P_gen[g,n]) == 0.0 && trunc(Int,value(P_gen[g,n])) != l
            count5 += 1
        end
        if value(P_gen[g,n]) < l && value(P_gen[g,n]) != 0.0
            count6 += 1
        end
    end
    println("number of total generators using their total capacity ", count4)
    println("number of generators using none of their total capacity ", count5)
    println("number of generators using some of their total capacity ", count6)
end
# println(gen_running)
