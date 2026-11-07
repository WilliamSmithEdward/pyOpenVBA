"""The members no live probe touches, for a probe that walks a whole class and must leave them out by name.

A probe's own code may not name them (tests/test_probes_leave_mail_alone.py, tests/test_probes_leave_crypto_alone.py), so a probe that builds
its list from docs/excel_checklist.csv takes them away with these sets instead. This module is not a probe and runs nothing.
"""
from __future__ import annotations

#: Could send mail from the account Excel runs under, or open a mail envelope.
MAIL = frozenset({'Mailer', 'EnvelopeVisible', 'MailEnvelope', 'SendMail', 'SendForReview', 'SendMailer', 'ForwardMailer', 'SendFaxOverInternet',
                  'RoutingSlip', 'HasRoutingSlip', 'HasMailer'})
#: Rights, signing, sensitivity labels and the work identity: a provider or a rights server can reach hardware or the network.
RIGHTS = frozenset({'Permission', 'Signatures', 'SensitivityLabel', 'WorkIdentity', 'SetPasswordEncryptionOptions'})
#: Servers, shared spaces and panes a workbook on a local disk has none of.
SERVERS = frozenset({'ServerPolicy', 'SharedWorkspace', 'DocumentLibraryVersions', 'ContentTypeProperties', 'Research', 'SmartDocument'})
