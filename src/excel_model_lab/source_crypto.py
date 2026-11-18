"""MS-XLDM CryptKey SIMPLEBLOB and connection-string envelope.

This is Excel's reversible storage format, not protection for secrets. The key
travels in the workbook. Never use this module as a credential vault.
Format reference: MS-XLDM 2.1.2.4; envelopes also checked against PivotWright's
MIT-licensed ModelCryptKey format implementation and native Excel samples.
"""
import base64
import os
import struct
from uuid import UUID
from .abf import UnsupportedModelError


def session_key(data):
    magic = UUID('5d21bc98-8d2d-4ee6-a8e5-d038aac94441').bytes_le
    if len(data) != 144 or data[:16] != magic or data[128:] != magic:
        raise UnsupportedModelError('Unsupported model CryptKey envelope')
    header_size, size = struct.unpack_from('<II', data, 20)
    if (header_size, size) != (48, 80) or data[48:50] != b'\x01\x02' or struct.unpack_from('<I', data, 52)[0] != 0x6603:
        raise UnsupportedModelError('Native source authoring requires a 3DES CryptKey')
    block = data[60:124][::-1]
    if block[:2] != b'\x00\x02':
        raise UnsupportedModelError('Unsupported CryptKey session-key block')
    separator = block.find(b'\0', 2)
    if separator < 10 or len(block[separator + 1:]) != 24:
        raise UnsupportedModelError('Invalid CryptKey session-key size')
    return block[separator + 1:]


def encrypt_string(key_file, text, *, flags=0x08000001):
    try:
        from cryptography.hazmat.decrepit.ciphers.algorithms import TripleDES
        from cryptography.hazmat.primitives.ciphers import Cipher, modes
        from cryptography.hazmat.primitives.padding import PKCS7
    except ImportError as exc:
        raise ImportError('OLE DB bindings require excel-data-model-lab[native-sources]') from exc
    if '\0' in text:
        raise ValueError('Source strings cannot contain NUL characters')
    raw = os.urandom(8) + (text + '\0').encode('utf-16le')
    padder = PKCS7(64).padder()
    padded = padder.update(raw) + padder.finalize()
    encryptor = Cipher(TripleDES(session_key(key_file)), modes.CBC(b'\0' * 8)).encryptor()
    encrypted = encryptor.update(padded) + encryptor.finalize()
    encoded = base64.b64encode(struct.pack('<II', 8, flags) + encrypted).decode('ascii')
    return ''.join(encoded[start:start + 76] + ' \n' for start in range(0, len(encoded), 76))
