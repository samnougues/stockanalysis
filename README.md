# Stockanalysis

## Overview of Project

### Purpose

The purpose of this analysis was to mimic real world issues with fixing someone else's code so that it runs faster.

## Results

- Most of the stocks in 2017  had a positive performance. As seen in this picture (https://github.com/samnougues/stockanalysis/blob/2ee446cb1cb3bb21f4b7d50d1c5bb89e48359fac/VBA_Challenge_2017.png), TERP was the only stock with a negative return. On the other hand, in 2018, most of the stocks had a negative return. As seen in this picture https://github.com/samnougues/stockanalysis/blob/0f415119fee62a10ca85de1bcc0f781f3897b7e2/VBA_Challenge_2018.png, the only stocks to have a positive return were ENPH and RUN. 

- The times for the results of the refactored code were much better than that of the original code. This time difference occurred because the data refers to the data that precedes it, making it easier to execute. For example, the old code did not declare Dim tickerVolumes(12) As Long. The old code also switches to the j variable, which may slow execution. 

## Summary

- What are the advantages or disadvantages of refactoring code?

The advantages to refactoring code are that it eliminates redundancy or excessive code, makes the code easier to read, and can make the code run faster. One disadvantage is the time taken to refactor the code may not be worth the time saved from refactoring. Also, it can create new bugs in the code if done improperly.

- How do these pros and cons apply to refactoring the original VBA script?


Refactoring the original VBA script made the code run faster, but the time spent refactoring was not worth extra milliseconds. There were many bugs I had to fix in the process.
