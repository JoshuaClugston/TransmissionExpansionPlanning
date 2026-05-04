This repo explores the Transmission Expansion Planning Problem using IEEE-RTS-96 data acquired from [this publically available, online source](https://www2.ee.washington.edu/research/pstca/rts/pg_tcarts.html). 

Within this repo, the following, primary components are provided: 
- Several .txt files including the original data acquired from the previously listed online resource
- Modified data based on the IEEE-RTS-96 data from the source listed above in an .xlsx file
- Mathematical model encoded using Julia's JuMP, and a Julia script to solve the JuMP model using various solvers such as HiGHS and CPLEX
- Python script to generate figures for exploration of results obtained from the analysis of the computational model's solution and construction

Discussion and analysis pertaining to the code written in this repo can be found [on my blog](https://joshuaclugston.com//2026/03/10/Transmission-Expansion-Planning-Problem).
