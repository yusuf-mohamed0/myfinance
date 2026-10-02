# -*- coding: utf-8 -*-
"""
SIGN_TEMPLATE_MANIFEST - sign the integrity manifest for this project.

Hashes every shipped source file (04_Source/*.py, site_template.html, art/*) and
writes 03_System/source_manifest.json, then signs it with an RSA-2048 key
(OpenSSL-style PEM in 03_System/.manifest_key). verify_system.py checks that the
recorded hashes still match the files on disk, and that the signature is valid.

Run after you edit any source file:
    python 04_Source/sign_template_manifest.py

If no key exists yet, one is generated (needs `cryptography` or `openssl`).
The private key stays local — never commit it.
"""
import glob
import hashlib
import json
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "04_Source")
SYS = os.path.join(BASE, "03_System")
KEY = os.path.join(SYS, ".manifest_key")
KEY_PUB = os.path.join(SYS, ".manifest_key.pub")
OUT = os.path.join(SYS, "source_manifest.json")
BANNER = "# rsa-2048"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def shipped_files():
    out = []
    for p in sorted(glob.glob(os.path.join(SRC, "*.py"))):
        out.append(os.path.relpath(p, BASE).replace("\\", "/"))
    for name in ("site_template.html",):
        p = os.path.join(SRC, name)
        if os.path.isfile(p):
            out.append(os.path.relpath(p, BASE).replace("\\", "/"))
    for p in sorted(glob.glob(os.path.join(SRC, "art", "*.png"))):
        out.append(os.path.relpath(p, BASE).replace("\\", "/"))
    return out


def ensure_key():
    """Return a private-key object, generating one on first run."""
    if os.path.isfile(KEY):
        try:
            from cryptography.hazmat.primitives import serialization as _ser
            with open(KEY, "rb") as fh:
                return _ser.load_pem_private_key(fh.read(), password=None)
        except Exception as exc:
            print("ERROR: cannot load {} ({})".format(KEY, exc))
            return None
    try:
        from cryptography.hazmat.primitives.asymmetric import rsa as _rsa
        from cryptography.hazmat.primitives import serialization as _ser
        key = _rsa.generate_private_key(public_exponent=65537, key_size=2048)
        pem = key.private_bytes(
            encoding=_ser.Encoding.PEM,
            format=_ser.PrivateFormat.PKCS8,
            encryption_algorithm=_ser.NoEncryption())
        with open(KEY, "wb") as fh:
            fh.write(pem)
        pub = key.public_key().public_bytes(
            encoding=_ser.Encoding.PEM,
            format=_ser.PublicFormat.SubjectPublicKeyInfo)
        with open(KEY_PUB, "wb") as fh:
            fh.write(pub)
        print("generated new RSA-2048 key at {}".format(os.path.relpath(KEY, BASE)))
        return key
    except Exception as exc:
        print("ERROR: cannot generate key ({}). Run: pip install cryptography".format(exc))
        return None


def sign(key, payload_b64):
    """Raw RSA-2048 PKCS#1 v1.5 over SHA-256 of the payload bytes."""
    import base64
    from cryptography.hazmat.primitives import hashes as _h
    from cryptography.hazmat.primitives.asymmetric import padding as _pad
    sig = key.sign(base64.b64decode(payload_b64),
                   _pad.PKCS1v15(), _h.SHA256())
    return base64.b64encode(sig).decode("ascii")



# Arabic output must not crash a cp1252 Windows console
def _mf_utf8():
    import sys as _s
    for _n in ("stdout", "stderr"):
        _st = getattr(_s, _n, None)
        if _st is not None:
            try:
                _st.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
_mf_utf8()
def main():
    key = ensure_key()
    if key is None:
        sys.exit(1)
    files = shipped_files()
    if not files:
        print("ERROR: no source files found under 04_Source")
        sys.exit(1)
    entries = {}
    for rel in files:
        entries[rel] = sha256(os.path.join(BASE, rel.replace("/", os.sep)))
    payload = {
        "algo": "sha256",
        "sig_algo": "rsa-2048/sha256",
        "count": len(entries),
        "files": entries,
    }
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False)
    import base64
    payload_b64 = base64.b64encode(body.encode("utf-8")).decode("ascii")
    sig = sign(key, payload_b64)
    doc = dict(payload)
    doc["signature"] = sig
    doc["payload_b64"] = payload_b64
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=2)
    print("manifest OK: {} files hashed + RSA-2048 signed".format(len(entries)))
    print("wrote {}".format(os.path.relpath(OUT, BASE)))
    print("private key stays local: {}".format(os.path.relpath(KEY, BASE)))


if __name__ == "__main__":
    main()