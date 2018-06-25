# CPET
A semi-generalized maximal oxygen consumption calculation

This is a short script that calculates various parameters of a cardio-pulmonal exercise test/maximal oxygen consumption test (VO2max test) performed on a cosmed device. 

Data from the device is output in .xlsx or .xls files depending on the device version. The output file consists of two data tables within the same spreadsheet, one for information on test parameters and participant pii, and another which contains individual measurements/observations. Observations are made breath-by-breath and identified by the elapsed test time. There may be more than one measurement per second and observations are not made at any uniform times, only when a breath is taken. A test typically runs app. 15min. 
As oxygen consumption is highly variable, simply finding the max value of each variable is insufficient. Rather, the highest running average over e.g. 20 seconds is calculated. 

This script

A) Aggregates multiple observations within the same second

B) Expands the observations to encompass all elapsed seconds with NA values at observation times where no measurement was made

C) Calculates running averages of variables of interest

D) Writes maximum running averages to a new destination

