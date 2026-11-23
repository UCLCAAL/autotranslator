#Do not use encrypted_keys() unless have own API key that want to encode as this will overwrite the existing keys.enc file. 
# Use load_keys() in main.py to load the keys for translation.
# keys.enc is encrypted with a team passphrase and is committed to GitHub. 
# The local .env file is git-ignored and stores the decrypted keys for convenience on your machine.
import base64
import getpass
import os
from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

def encrypted_keys():
    pw1 = getpass.getpass("Create shared team passphrase: ")
    pw2 = getpass.getpass("Confirm shared team passphrase: ")
    if not pw1 or pw1 != pw2:
        raise ValueError("Passphrases were empty or did not match.")

    raw_keys = input("Paste comma-separated DeepL API keys: ").strip()
    if not raw_keys:
        raise ValueError("No API keys entered.")

    salt = os.urandom(16)
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=480000,
    )
    derived_key = base64.urlsafe_b64encode(kdf.derive(pw1.encode("utf-8")))
    encrypted_token = Fernet(derived_key).encrypt(raw_keys.encode("utf-8"))

    with open("keys.enc", "wb") as f:
        f.write(salt + encrypted_token)

    # Also save locally to .env so you don't need to re-enter the passphrase yourself
    with open(".env", "w", encoding="utf-8") as f:
        f.write(raw_keys)

    print("Saved keys.enc (commit to GitHub) and local .env (ignored by Git).")

if __name__ == "__main__":
    encrypted_keys()

def load_keys(passphrase=None):
    # 1. Read from local git-ignored .env if already unlocked on this machine
    if os.path.exists(".env"):
        with open(".env", "r", encoding="utf-8") as f:
            keys = [k.strip() for k in f.read().split(",") if k.strip()]
            if keys:
                return keys

    # 2. Otherwise decrypt keys.enc using the passphrase (passed as arg or prompted)
    if not os.path.exists("keys.enc"):
        raise FileNotFoundError("keys.enc not found. Run encrypt_keys.py first.")

    if not passphrase:
        passphrase = getpass.getpass("Enter team passphrase to unlock DeepL keys: ")

    with open("keys.enc", "rb") as f:
        data = f.read()

    salt, token = data[:16], data[16:]
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=480000)
    derived_key = base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))

    try:
        decrypted = Fernet(derived_key).decrypt(token).decode("utf-8")
    except InvalidToken:
        raise ValueError("Incorrect passphrase. Could not decrypt keys.enc.")

    with open(".env", "w", encoding="utf-8") as f:
        f.write(decrypted)

    return [k.strip() for k in decrypted.split(",") if k.strip()]