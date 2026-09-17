# Model Formulation

## 1. Problem Type

The problem addressed is a **product-mix determination problem**: given a fixed portfolio of parts, each requiring processing time on one or more of several shared machine groups with finite capacity, determine the production quantity of each part that maximizes total economic contribution. This is a classical constrained profit-maximization **Linear Programming (LP)** problem.

## 2. Notation

| Symbol | Meaning |
|---|---|
| $i = 1, \dots, n$ | Index over parts in the portfolio |
| $m = 1, \dots, M$ | Index over machine groups |
| $x_i$ | Decision variable — quantity of part $i$ to produce |
| $p_i$ | Unit selling price of part $i$ |
| $t_{i,m}$ | Processing time (hours) required by one unit of part $i$ on machine group $m$ |
| $c_m$ | Hourly operating-cost coefficient of machine group $m$ |
| $Cap_m$ | Available capacity (hours) of machine group $m$ over the planning period |
| $x_i^{min}$ | Minimum production floor applied to part $i$ (placeholder value — see Section 7) |

## 3. Objective Function

The objective is to maximize total economic contribution, defined as total revenue minus total machine-usage cost:

$$\max Z = \underbrace{\sum_{i=1}^{n} p_i x_i}_{\text{Total Revenue}} - \underbrace{\sum_{m=1}^{M} c_m \left(\sum_{i=1}^{n} t_{i,m} x_i\right)}_{\text{Total Machine Usage Cost}}$$

The inner sum $\sum_i t_{i,m} x_i$ represents the total hours consumed on machine group $m$ given the current production plan; multiplying by $c_m$ converts this into a cost figure, which is subtracted from revenue.

## 4. Constraints

**Capacity constraint** — total hours consumed on each machine group cannot exceed its available capacity:

$$\sum_{i=1}^{n} t_{i,m} x_i \le Cap_m \qquad \forall m = 1, \dots, M$$

**Minimum production floor** — production must meet a floor value entered into the model to ensure a feasible, non-trivial solution:

$$x_i \ge x_i^{min} \qquad \forall i = 1, \dots, n$$

**Integrality constraint** — production quantities must be whole units:

$$x_i \in \mathbb{Z}^{+} \qquad \forall i = 1, \dots, n$$

## 5. Why Linear Programming

The relationship between the decision variables ($x_i$) and both the objective (revenue and cost) and the constraints (machine-hours consumed) is linear: doubling $x_i$ exactly doubles both its revenue contribution and its machine-hour consumption, for every part and every machine group. This linearity is what makes LP — rather than a nonlinear or heuristic approach — the structurally appropriate technique here, and it is also why the Simplex-based Excel Solver engine can solve the model to (near-)optimality efficiently even with hundreds of decision variables.

## 6. Solver Implementation Notes

- The objective cell, decision variable range, and constraints are set programmatically via VBA (`SolverOk`, `SolverAdd`, `SolverSolve`) rather than manually through the Solver dialog, so the model can be re-solved automatically whenever input data changes.
- Capacity constraints are added **one at a time per machine group** rather than as a single multi-range constraint, since Excel Solver can raise errors when a single constraint call spans mismatched ranges.
- Before solving, the macro validates that every priced part has at least one machine-hour entry — a part with a price but no assigned machine time would make the model unbounded (infinite profit), since it could be "produced" without consuming any resource.

## 7. Limitations of the Current Parameterization

- The cost coefficients $c_m$ are currently preliminary estimates rather than the output of a formal, machine-level cost study (depreciation, energy, tooling, maintenance).
- Machine groups that physically consist of multiple parallel units are modeled with a single, uniform $Cap_m$ value rather than a capacity scaled by the number of machines in the group.
- Processing times $t_{i,m}$ for some parts have not yet been confirmed through formal time studies.
- The minimum production floor $x_i^{min}$ applied to every part (2 units in the current run) is **not a validated customer commitment**; it is a floor value entered to guarantee a feasible, non-trivial solution, and could equally be set to a different figure (e.g., 1 or 3) without changing the structure of the model. As a direct consequence, any observed pattern of machine groups reaching 100% utilization is an artifact of this placeholder value combined with the uniform capacity assumption above — it should not be read as a validated bottleneck finding.
