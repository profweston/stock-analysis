# VBA Challenge using Analysis of Stock Data

## Overview

This project analyzes data for 12 stocks related to Green Energy Companies to calculate their returns during the years 2017 and 2018. A VBA script utilizing nested loops was developed to find total volume and starting and ending price to produce a return percent. Next the same analysis was done after refactoring the VBA code to remove the nested loop design and implement an index design. A summary of the stocks performance as well as execution performance of the VBA scripts is presented here.

## Results


### Stock Performance



![Stock Performance](/Resources/Stock_Performance.png)


### Execution Performance

![Timer 2017](/Resources/Timer_2017_Nest_Loop.png)

## Summary


### Advantages of Refactoring

**Refactoring** refers to the process of restructing computer code without changing its functionality [@https://en.wikipedia.org/wiki/Code_refactoring]. The most obvious advantage of refactoring is by improving the structure to be more straightforward, it naturally makes it easier to work with. Consequently, int he case of bugs, they can be identified more readily and remedied. If the code is written in a design pattern, then it could be adapted to another project or easily extended to the current project to do deeper analyses. Moreover, if the code is written in an efficient manner it could reduce the processing time. [@https://anarsolutions.com/code-refactoring-concept-analysis/]

### Disadvantages of Refactoring

While refactoring can offer advantages, it also has a down side. Refactoring can be very time-consuming and could potentially introduce bugs into a script that was working previously. [@https://anarsolutions.com/code-refactoring-concept-analysis/]
