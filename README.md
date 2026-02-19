# CreateSimplifiedOFX

`CreateSimplifiedOFX.ps1` generates a simplified OFX 1.02 (SGML) file containing annual `CREDIT` transactions for import into personal-finance tools.

## Inputs

Pass one or more `-Entry` values, each as:

`MM/DD/YY,Amount,Years`

Where:
- `MM/DD/YY` is normalized to four-digit year using a `20` prefix (for example, `2/1/30` becomes `02/01/2030`).
- `Amount` is a decimal value using `.` as the separator (for example, `21.78`).
- `Years` is a non-negative integer and is **inclusive** (`Years + 1` transactions are generated).

## Example

```powershell
./CreateSimplifiedOFX.ps1 \
  -Entry "08/15/26,21.38,29" \
  -Entry "02/15/27,15.00,5" \
  -OutputPath "./annual_deposits.ofx"
```

## Output details

- Account metadata is fixed to:
  - `<ACCTID>0`
  - `<ACCTTYPE>IMPORT`
- Transaction type is always `CREDIT`.
- Output charset is UTF-8 with BOM (`ENCODING:UTF-8`, `CHARSET:65001`).
- Amount formatting is invariant culture so OFX amounts consistently use `.`.


## Script files

- `CreateSimplifiedOFX.ps1` is the primary implementation.
- `createSimplifiedOFX.ps` is a compatibility wrapper that forwards arguments to `CreateSimplifiedOFX.ps1` so both entry points behave the same way.
- Default output path is `C:\output\import.ofx`.
- Writes to the Windows system directory are blocked by both CLI and UI validation.
