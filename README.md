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

For one series, from PowerShell 7:

```powershell
pwsh -NoProfile -File ./createSimplifiedOFX.ps -Entry "08/15/26,21.38,29" -OutputPath "./annual_deposits.ofx"
```

For multiple series, run this **inside PowerShell**. The script retains its original `.ps` filename, so use a script block to load it explicitly. Pass an array to `-Entry` once; repeating the parameter is not supported by PowerShell.

```powershell
& ([scriptblock]::Create((Get-Content -Raw ./createSimplifiedOFX.ps))) `
  -Entry @("08/15/26,21.38,29", "02/15/27,15.00,5") `
  -AccountType CHECKING `
  -OutputPath "./annual_deposits.ofx"
```

## Output details

- Account metadata is fixed to:
  - `<ACCTID>0`
  - `<ACCTTYPE>CHECKING` by default. Set `-AccountType` to `CHECKING`, `SAVINGS`, `MONEYMRKT`, or `CREDITLINE` to match the destination account.
- Transaction type is always `CREDIT`.
- Output charset is Windows-1252 (`CHARSET:1252`).
- Amount formatting is invariant culture so OFX amounts consistently use `.`.
- Dates use the Gregorian calendar regardless of the system locale.
- `DTSERVER` is the current UTC generation time, rather than the final scheduled deposit date.
- Files use Windows CRLF line endings, with a blank line separating the header and SGML body, and no byte-order mark.
- Bank ID `000000000`, account ID `0`, and ledger balance `0.00` are synthetic placeholders, not a real bank statement or account balance.

## Microsoft Money format review

The original `<ACCTTYPE>IMPORT` is outside the OFX bank account type enumeration. The independent `ofxtools` validator rejected the original output with `Can't set BANKACCTFROM.accttype to IMPORT`. The default is now `CHECKING`, and unsupported account types are rejected before writing a file.

The original decimal parsing and `TRNAMT` formatting depended on the system culture. This could reject dot-decimal input or write comma-decimal amounts; calendars such as the Thai Buddhist calendar could also change the transaction year. Parsing and output now explicitly use invariant culture and Gregorian dates. The memo's dollar amount was also missing because PowerShell expanded `${0:N2}` as a variable; its format string is now literal.

The existing OFX 1.x header (`OFXHEADER:100`, `DATA:OFXSGML`, `VERSION:102`, `ENCODING:USASCII`, `CHARSET:1252`), omitted leaf closing tags, and explicitly closed aggregate tags are consistent with SGML OFX. They do not need conversion to XML. CRLF line endings are a conservative choice for the Windows importer; LF alone was not established as the cause of Money's error.

References used for the review:

- [OFX 1.x DTD maintained by libofx](https://github.com/libofx/libofx/blob/ea5d1cb71a6319b689a87249d53c96f719821113/dtd/ofx160.dtd): the CheckFree/Intuit/Microsoft DTD's `ACCOUNTENUM` and bank statement structure. This is the later 1.6 DTD, used to cross-check the shared bank statement fields, not a complete 1.02 certification.
- [ofxtools bank statement models](https://github.com/csingley/ofxtools/blob/master/ofxtools/models/bank/stmt.py): bank account enumeration, required statement fields, and transaction types; see the OFX section references in the model documentation.

The official OFX specification download hosts were inaccessible from the review environment. Corrected files pass independent parsing and model validation, but Microsoft Money itself is unavailable here. Regenerate your file and retry the import in Money; this validation cannot guarantee every Money version's behavior, particularly for synthetic account metadata and future-dated transactions.

## Regression checks

Python is needed only for tests; generating OFX files still needs only PowerShell.

```sh
python3 -m venv /tmp/ofx-tests
/tmp/ofx-tests/bin/pip install -r tests/requirements.txt
/tmp/ofx-tests/bin/python -m unittest discover -s tests -v
```

Put `pwsh` on `PATH`, or set `PWSH` to its executable path. Tests exercise the real script and parse every successful output with `ofxtools`, including multiple series, all supported account types, zero additional years, leap days, five regional settings, and invalid inputs.
