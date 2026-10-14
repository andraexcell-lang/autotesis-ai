import os
from cryptography import x509
from cryptography.hazmat.primitives import serialization

backend_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(backend_dir, "localhost.crt"), "rb") as f:
    cert = x509.load_pem_x509_certificate(f.read())

with open(os.path.join(backend_dir, "localhost.cer"), "wb") as f:
    f.write(cert.public_bytes(serialization.Encoding.DER))

with open(os.path.join(backend_dir, "ca.crt"), "rb") as f:
    ca = x509.load_pem_x509_certificate(f.read())

with open(os.path.join(backend_dir, "ca.cer"), "wb") as f:
    f.write(ca.public_bytes(serialization.Encoding.DER))

print("Created localhost.cer and ca.cer (DER format)")
