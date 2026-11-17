"""End-to-end OFX checks using the independent ofxtools parser/validator."""
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

from ofxtools.Parser import OFXTree

SCRIPT = Path(__file__).resolve().parents[1] / 'createSimplifiedOFX.ps'
PWSH = os.environ.get('PWSH') or shutil.which('pwsh')


class OfxTests(unittest.TestCase):
    def generate(self, entries, culture='en-US', account_type='CHECKING', success=True):
        self.assertTrue(PWSH, 'Install PowerShell or set PWSH to its executable path')
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'deposits.ofx'
            # JSON transports test data without shell/PowerShell interpolation.
            config = Path(directory) / 'input.json'
            config.write_text(json.dumps(dict(script=str(SCRIPT), output=str(output),
                                             entries=entries, culture=culture,
                                             account_type=account_type)))
            runner = Path(directory) / 'run.ps1'
            runner.write_text('''param([string]$Config)
$ErrorActionPreference = 'Stop'
$c = Get-Content -Raw $Config | ConvertFrom-Json
[System.Threading.Thread]::CurrentThread.CurrentCulture = [cultureinfo]$c.culture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [cultureinfo]$c.culture
& ([scriptblock]::Create((Get-Content -Raw $c.script))) -Entry $c.entries -OutputPath $c.output -AccountType $c.account_type
''')
            result = subprocess.run([PWSH, '-NoLogo', '-NoProfile', '-File', str(runner),
                                     str(config)], capture_output=True, text=True)
            if not success:
                self.assertNotEqual(result.returncode, 0, result.stdout)
                self.assertFalse(output.exists(), 'Invalid input must not write an OFX file')
                return
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            raw = output.read_bytes()
            tree = OFXTree()
            tree.parse(str(output))
            model = tree.convert()  # Validates required aggregates, enums and data types.
            return raw, model

    def test_documented_multiple_series(self):
        raw, model = self.generate(['08/15/26,21.38,29', '02/15/27,15.00,5'])
        text = raw.decode('cp1252')
        self.assertTrue(raw.startswith(b'OFXHEADER:100\r\n'))
        self.assertIn(b'NEWFILEUID:NONE\r\n\r\n<OFX>', raw)
        self.assertNotIn(b'\n', raw.replace(b'\r\n', b''))
        self.assertIn('VERSION:102', text)
        self.assertIn('CHARSET:1252', text)
        self.assertEqual(len(model.statements), 1)
        statement = model.statements[0]
        self.assertEqual(statement.bankacctfrom.accttype, 'CHECKING')
        self.assertEqual(statement.bankacctfrom.acctid, '0')
        transactions = statement.banktranlist
        self.assertEqual(len(transactions), 36)
        self.assertEqual(len({t.fitid for t in transactions}), 36)
        self.assertTrue(all(t.trntype == 'CREDIT' for t in transactions))
        self.assertEqual(str(transactions[0].trnamt), '21.38')
        self.assertEqual(transactions[0].dtposted.strftime('%Y%m%d'), '20260815')
        self.assertEqual(transactions[-1].dtposted.strftime('%Y%m%d'), '20550815')
        self.assertEqual([t.dtposted for t in transactions], sorted(t.dtposted for t in transactions))
        self.assertRegex(text, r'<DTSERVER>\d{14}\[0:GMT\]')

    def test_regional_settings_and_calendars(self):
        for culture in ['en-US', 'de-DE', 'fr-FR', 'th-TH', 'ar-SA']:
            with self.subTest(culture=culture):
                raw, _ = self.generate(['2/1/30,21.78,2'], culture=culture)
                text = raw.decode('cp1252')
                self.assertEqual(re.findall(r'<TRNAMT>([^\r\n]+)', text), ['21.78'] * 3)
                self.assertEqual(re.findall(r'<DTPOSTED>([^\r\n]+)', text),
                                 ['20300201', '20310201', '20320201'])
                self.assertIn('<MEMO>Deposit $21.78 on 02/01/2030', text)

    def test_zero_years_and_account_types(self):
        for kind in ['CHECKING', 'SAVINGS', 'MONEYMRKT', 'CREDITLINE', 'savings']:
            with self.subTest(kind=kind):
                _, model = self.generate(['2/1/30,21.78,0'], account_type=kind)
                statement = model.statements[0]
                self.assertEqual(statement.bankacctfrom.accttype, kind.upper())
                self.assertEqual(len(statement.banktranlist), 1)

    def test_leap_day(self):
        _, model = self.generate(['2/29/28,1000.50,1'])
        self.assertEqual([t.dtposted.strftime('%Y%m%d') for t in model.statements[0].banktranlist],
                         ['20280229', '20290228'])

    def test_invalid_inputs(self):
        for entry in ['2/30/26,1,2', '2/1/26,1,-1', 'malformed', '2/1/26,1.2.3,2']:
            with self.subTest(entry=entry):
                self.generate([entry], success=False)
        self.generate(['2/1/26,1,2'], account_type='IMPORT', success=False)


if __name__ == '__main__':
    unittest.main()
