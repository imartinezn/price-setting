# Price-setting technologies in EU electricity markets
---

## Introduction

Here we present the code behind the manuscript "Price-setting tipping points in decarbonising EU electricity markets" (in preparation), an open source framework for reverse-engineering price-setting technologies based on observed generation, cross-border flows, wholesale prices, and commodity costs. This framework is used to assess a decade of hourly data across EU bidding zones and capture technological marginality across the full generation spectrum, from variable renewables (VRE, i.e. solar and wind power) and hydropower to nuclear and fossil-powered units. 
The gained empirical insights can inform energy policy across multiple domains: clarifying the coupling of electricity prices and fossil fuel costs, quantifying the effect of increased VRE generation in displacing fossil fuel price-setters, and illustrating how interconnections affect price-setting.

## Diagram
Below, a schematic diagram of the logic of the price-setting identifying algorithm is displayed.

# <img src="./diagram/FigSupplA.png" align="right" />

## Repository Structure

This price-setting framework comprises one executable python script, under the control of an Excel-based "control_file" and two more Excel files to support the process of the user's commodity and price, generation, flows, loads and reservoir level data.
The repository contains the following files:
* **Price_setting_publishable.py** (the actual code)
* **control_file.xlsx** (providing crucial input parameters to the code)
* **user_commodity_file.xlsx** (providing commodity time series to the code)
* **example_inputs.xlsx** (providing generation and flow time series to the code; optional)

## License
See also the [LICENSE](./LICENSE.md) file.

## Authors

Ignacio Martínez, Directorate-General for Energy, European Commission, ignacio.martinez1@ec.europa.eu

Derck Koolen, Directorate-General for Energy, European Commission & School of Economics, Utrecht University, derck.koolen@ec.europa.eu

Massimo Tognoni, Directorate-General for Energy, European Commission, massimo.tognoni@ec.europa.eu

Sebastian Sterl, Directorate-General for Energy, European Commission & Faculty of Engineering, Vrije Universiteit Brussel (VUB), sebastian.sterl@vub.be
