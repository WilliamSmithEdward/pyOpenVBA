"""Workbook.SetPasswordEncryptionOptions, as live Excel answered it once: retired, and not to be run again.

On 2026-11-07 this script called SetPasswordEncryptionOptions on new workbooks, held in memory and never saved, 123 cases bound typed and
late. It named ten Windows cryptographic providers with nine algorithm names each, among them "Microsoft Base Smart Card Crypto Provider".
Naming that provider raised Windows Security prompts on the machine asking for a smartcard. After each call it read PasswordEncryptionProvider,
PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength and PasswordEncryptionFileProperties.

tests/fixtures/workbook_encryption_options.json keeps the cases and what Excel answered: do not remeasure them. No live probe names a
provider, or touches encryption options, rights, signatures, certificates or sensitivity labels (tests/test_probes_leave_crypto_alone.py),
so the cases live in the fixture, not in this script's code, and the script refuses to run.
"""
from __future__ import annotations

import sys


def main() -> None:
    sys.exit('retired: this probe reached Windows CryptoAPI and must not run; tests/fixtures/workbook_encryption_options.json keeps its answers')


if __name__ == '__main__':
    main()
