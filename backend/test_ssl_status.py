import ssl
import urllib.request
import os

ca_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ca.crt")
cert_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "localhost.crt")
key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "localhost.key")

print("Files exist:")
print("ca.crt:", os.path.exists(ca_path))
print("localhost.crt:", os.path.exists(cert_path))
print("localhost.key:", os.path.exists(key_path))
