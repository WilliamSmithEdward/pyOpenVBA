# The ClamAV engine the malware scan runs, pinned by digest;
# update-images.yml proposes new releases and rebuilds. The image carries a
# copy of the signatures, but freshclam brings them up to date on every
# run, and the report records the database version it scanned with.
FROM clamav/clamav:1.5.4@sha256:7054e3f34308c384ea43cd786ea17abf899bd44af40af9f1b099853d7a5c214e
