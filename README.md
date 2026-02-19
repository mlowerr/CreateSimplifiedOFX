# CreateSimplifiedOFX

`createSimplifiedOFX.ps` generates a simplified OFX 1.02 (SGML) file containing annual `CREDIT` transactions for import into personal-finance tools.

## Inputs

Pass one or more `-Entry` values, each as:

`MM/DD/YY,Amount,Years`

Where:
- `MM/DD/YY` is normalized to four-digit year using a `20` prefix (for example, `2/1/30` becomes `02/01/2030`).
- `Amount` is a decimal value using `.` as the separator (for example, `21.78`).
- `Years` is a non-negative integer and is **inclusive** (`Years + 1` transactions are generated).

## Example

```powershell
./createSimplifiedOFX.ps \
  -Entry "08/15/26,21.38,29" \
  -Entry "02/15/27,15.00,5" \
  -OutputPath "./annual_deposits.ofx"
```

## Output details

- Account metadata is fixed to:
  - `<ACCTID>0`
  - `<ACCTTYPE>IMPORT`
- Transaction type is always `CREDIT`.
- Output charset is Windows-1252 (`CHARSET:1252`).
- Amount formatting is invariant culture so OFX amounts consistently use `.`.
