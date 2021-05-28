# CNC-Run-Time-Estimator
Gathers sheet material counts and shopbot (.sbp) files from the Solidworks PDM and runs a homemade CNC run time estimation algorithm. All this information can be used to input into cost analyses for Creative Works.

API's used: Solidworks 2020 API and Solidworks Professional PDM API

Dependent Factors:
File names in the SWPDM must adhere to naming standards set by CWI: {#}\_{Material} {thickness}{sheet number}.sbp
There must be a PROD-#### CNC file for the program to find. This will serve as a standard catch. The program will mention if it doesn't find a file by this name. If there isn't, there should be. This will give the user a chance to fix the error before running the program again.
