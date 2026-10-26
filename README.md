# price-setting
Open source framework for reverse-engineering price-setting technologies in the EU electricity markets

Price-setting tipping points in decarbonising EU electricity markets

Introduction:

Here we present the code behind the manuscript "Price-setting tipping points in decarbonising EU electricity markets", an open source framework for reverse-engineering price-setting technologies based on observed generation, cross-border flows, wholesale prices, and commodity costs. This framework is used to assess a decade of hourly data across EU bidding zones and capture technological marginality across the full generation spectrum, from variable renewables (VRE, i.e. solar and wind power) and hydropower to nuclear and fossil-powered units. 
The gained empirical insights can inform energy policy across multiple domains:  clarifying the coupling of electricity prices and fossil fuel costs, quantifying the effect of increased VRE generation in displacing fossil fuel price-setters, and illustrating how interconnections affect price-setting.

Price-setting algorithm diagram:
Schematic diagram of the logic of the price-setting identifying algorithm, as described in Methods section of the manuscript.

FigSuppIA.eps

Repository structure:

This price-setting framework comprises one python script, under the control of an excel "control_file" and two more excel files to support the process of the user's commodity and price, generation, flows, loads and reservoir level data.
The repository contains the following files:
- Price_setting_publishable.py
- control_file.xlsx
- user_commodity_file.xlsx
- example_inputs.xlsx


Authors
Ignacio Martinez, Directorate-General for Energy, European Commission, ignacio.martinez1@ec.europa.eu

Derck Koolen, Directorate-General for Energy, European Commission & School of Economics, Utrecht University, derck.koolen@ec.europa.eu

Massimo Tognoni, Directorate-General for Energy, European Commission, massimo.tognoni@ec.europa.eu

Sebastian Sterl, Directorate-General for Energy, European Commission & Vrije Universiteit Brussel (VUB), Sebastian.Sterl@vub.be
