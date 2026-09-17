## Overview

This repository contains the VBA/Solver-based optimization model developed during my industrial engineering internship at a job-shop CNC manufacturing company. The company produces a wide variety of precision-machined parts in low volumes across 14 CNC machine groups, and needed a way to determine which parts — and in what quantities — should be produced to maximize total economic contribution without exceeding available machine capacity.

The model is a Linear Programming (LP) formulation, solved directly within Microsoft Excel using the Solver add-in, driven by a VBA macro that performs input validation and automates the Solver call. 

## How the VBA Macro Works

The macro (`vba/MaksimumKar_KesinCozum_V3.bas`) automates:

1. **Input validation** — checks that capacity values are numeric and non-negative, that every priced part has at least one machine-hour entry (otherwise the model would be unbounded), and that available capacity is sufficient for the minimum quantity of every part.
2. **Solver setup** — sets the objective cell (total economic contribution) to maximize.
3. **Constraint definition** — adds the minimum-quantity, integrality, and per-machine capacity constraints.
4. **Solve** — runs Solver in unattended mode and reports success/failure.

This removes the need to manually re-enter constraints in the Solver dialog every time input data changes.
