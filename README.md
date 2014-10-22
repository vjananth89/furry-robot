furry-robot
===========

A small piece of code for finding duplicates in two rows in MS Excel. Distinct values are printed in a new column.


###Why?

Very often we find ourselves filling up excels and find ourselves stare at huges lists of columns. furry-robot is a simple piece of code I wrote (probably as a note to self) to quickly compare two lists and print the unique values.

The code runs as a Macro on MS Excel versions. 

###How?

Lets say we have an excel that looks like the one below.

A        |   B
-------- | -------
SFO      |   JAX
LAX      |   IAH
JFK      |   ALB
EWR      |   SFO
IAH      |   JFK
ALB      |
JAX      |
MCI      |



Looking at the list, we know that both have same values, but we do not want to spend time unecessarily reading every cell, line by line.

The code prints the results as below


A     |    B    |  C 
----- |  ------ | ------
SFO   |   JAX   | LAX
LAX   |   IAH   | EWR
JFK   |   ALB   | MCI
EWR   |   SFO   |
IAH   |   JFK   |
ALB   |         |
JAX   |         |
MCI   |         |




>###NOTE:
>The macro given should be run in the **Visual Basic editor** that comes with MS Excel. It can be found in  **Tools -> Macro --> > Visual Basic Editor** 
>The code is found in Test.txt and can be copied into the Visual Basic Text editor
