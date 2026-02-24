import sys

import keyring

cred = keyring.get_credential("<REDACTED>", None)
if not cred or cred.password is None:
    print("No se encontraron credenciales en keyring.", file=sys.stderr)
    sys.exit(1)

print(cred.password)
