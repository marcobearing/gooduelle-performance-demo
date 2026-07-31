# Gooduelle Performance Demo

Public demonstration dashboard using fully synthetic operational data.

## Important

- Gooduelle is a fictional organization.
- All costs, budgets, FTE values, targets, histories and local benchmarks are synthetic.
- Benchmark values are illustrative and are not tied to any confidential source.
- No client workbook, identifier, comment or hidden metadata is included.

## Run locally

```cmd
npx --yes serve -l 5173 .
```

Open `http://127.0.0.1:5173/`.

The dashboard opens with built-in demo data. You can also import `Gooduelle_Performance_Demo.xlsx`.

## Traceability

The workbook behaves like a complete planning file:

- every function, legal entity, country and cost category has an explicit FY26/27 budget row;
- no hidden `current × 95%` budget fallback is used;
- monthly baseline values reconcile exactly to annual values;
- KPI definitions and monthly histories are included;
- each generated dataset has a stable Trace ID, generation method and public seed;
- `Data Dictionary` and `Generation Log` explain how every dataset was produced.
