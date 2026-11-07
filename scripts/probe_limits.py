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
#: Kept for the user, in the registry, across workbooks and sessions: a probe that sets one changes William's Excel. Only a probe that reads
#: it first and puts it back may set it (tests/test_probes_leave_user_settings_alone.py). DoNotPromptForConvert is DisplayConversionAlert.
USER_LEVEL = frozenset({'DoNotPromptForConvert'})
#: Never set by a probe: a mail envelope on the screen, saving to the cloud, a password a save encrypts or locks the file with, and the
#: provider it encrypts with, which names a cryptographic provider.
NEVER_SET = frozenset({'EnvelopeVisible', 'AutoSaveOn', 'Password', 'WritePassword', 'EncryptionProvider'})
#: Servers, shared spaces and panes a workbook on a local disk has none of.
SERVERS = frozenset({'ServerPolicy', 'SharedWorkspace', 'DocumentLibraryVersions', 'ContentTypeProperties', 'Research', 'SmartDocument'})
