# Concentration-Time Graphs for multi-inputs CSTR

## Background
[Continuous stirred-tank reactors (CSTR)](https://en.wikipedia.org/wiki/Continuous_stirred-tank_reactor) are common tools used in System Chemistry and Origins of Life research. A common laboratory setup includes a chamber with input(s), output, and a stirrer. Reagents are supplied by tubing from syringes, and the concentrations of reagents are controlled by adjusting the flow rate from each syringe. By analyzing the solution from the output, one can draw conclusions about the process happening inside the reactor.

[Example of CSTR for System Chemistry research from [*S.Semenov et al., Nature, 2016*](https://www.nature.com/articles/nature19776)](https://www.researchgate.net/profile/Jerome-Fox/publication/308706927/figure/fig2/AS:411560389431300@1475135051659/Schematic-representation-of-the-CSTR-experimental-set-up-In-this-set-up-syringe-pumps.png)

However, in some cases with complex systems, the concentrations of reagents (flow rates) need to be changed during the experiment. This program calculates the actual concentrations of reagents over time inside the multi-input CSTR, taking into account:
* CSTR parameters (size)
* Lag time caused by tubing
* Starting flow rates and concentrations of reagents
* Changes in flow rates during the experiments 


## Installation
Download files from this repository to your computer and install required Python packages with the command:
`pip install -r requirements.txt`

## Usage
To run the programm use command:
`python CSTR_graph.py`

The program has a user-friendly graphical interface. Fill the boxes in the opened window with all the parameters needed for the calculations:
* CSTR size
* Number of syringes
* Tubing length and diameter for each syringe
* Names of reagents in each syringe
* Concentrations of reagents in each syringe
* Starting flow rates for each syringe
* Changes in flow rates made during the experiment
* Exact time of changes
* Total time of the experiment
* Name of the experiment 

To start calculations, use the "Submit" button.
As the output, you will receive an Excel file with the concentration-time data as well as a graph for this data.

## Tests
To ensure that the program works correctly on your computer, you may run tests with the command:
`pytest`
 
