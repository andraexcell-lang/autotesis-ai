import os
from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption

backend_dir = os.path.dirname(os.path.abspath(__file__))
pfx_path = os.path.join(backend_dir, "server.pfx")
cert_path = os.path.join(backend_dir, "cert.pem")
key_path = os.path.join(backend_dir, "key.pem")

with open(pfx_path, "rb") as f:
    pfx_data = f.read()

private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
    pfx_data, b"autotesis123"
)

with open(cert_path, "wb") as f:
    f.write(certificate.public_bytes(Encoding.PEM))

with open(key_path, "wb") as f:
    f.write(private_key.private_bytes(
        encoding=Encoding.PEM,
        format=PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=NoEncryption()
    ))

print("Successfully extracted cert.pem and key.pem from server.pfx (TrustedPeople store)!")
