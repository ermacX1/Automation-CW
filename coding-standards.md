---
inclusion: auto
---

# Coding Standards — CW Automation

## General
- Use **Polars** for all data manipulation. Do not introduce pandas unless needed for openpyxl interop at the final write step.
- All DataFrames should use `infer_schema=False` on read (strings) then explicit `.cast()` where numerics are needed.
- Use type annotations on all function signatures.
- Keep functions pure where possible — accept DataFrames in, return DataFrames out.

## File I/O
- Raw feeds are **tab-delimited** (`\t`) text files. Always use `separator="\t"` and `truncate_ragged_lines=True`.
- CSVs (reference tables) are comma-delimited. Use default `pl.read_csv(...)`.
- Strip whitespace from all column names immediately after loading: `df.rename({c: c.strip() for c in df.columns})`.
- Skip files ≤ 400 bytes (header-only, no data rows).

## Naming Conventions
- Functions: `step<N>_<short_name>` for pipeline steps, `_helper_name` for internal helpers.
- Constants: `UPPER_SNAKE_CASE`.
- Variables: `lower_snake_case`.
- No single-letter variable names outside list comprehensions.

## Error Handling
- Use `strict=False` on all `.cast()` and `.str.to_date()` calls — bad data should become null, not crash the pipeline.
- Log warnings for missing reference data but continue processing.
- SFTP errors should fail loudly (raise) so the user knows the download didn't complete.

## Performance
- Prefer vectorised Polars expressions over `.map_elements()` / `.apply()`.
- Use `.map_elements()` only when no Polars expression exists (e.g., substring extraction with complex logic).
- Avoid collecting LazyFrames until the last possible moment if you refactor to lazy mode.
- Date padding (single-digit → zero-padded) uses vectorised `.str.replace()` — do not use `map_elements` for this.

## SFTP / Network
- Use **paramiko** for SFTP connectivity.
- PPK v2 RSA keys are parsed in-memory (no external `puttygen` dependency).
- Only download files newer than what's already on disk (compare file size to skip).
- Download the "latest batch" — files within 24h of the newest modification time on the server.

## Excel Output
- Use openpyxl for writing. Convert Polars → pandas only at the final write boundary (`final_df.to_pandas()`).
- Always preserve the template workbook's existing sheets — only overwrite Sheet2.
- Final output must have exactly 43 columns in the documented order.

## Product Filtering
- Only products in `VALID_CW_PRODUCTS` set are included in the pipeline.
- When a new CW product is added to the Access UNION ALL, add its 3-letter code to this set.
