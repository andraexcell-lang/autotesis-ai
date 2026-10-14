import os
import shutil

home = os.path.expanduser("~")
dev_certs_dir = os.path.join(home, ".office-addin-dev-certs")
print(f"Checking: {dev_certs_dir}")
if os.path.exists(dev_certs_dir):
    files = os.listdir(dev_certs_dir)
    print("Files in dev-certs:", files)
    
    # Copy localhost.crt and localhost.key to backend if they exist
    backend_dir = os.path.dirname(os.path.abspath(__file__))
    for f in files:
        src = os.path.join(dev_certs_dir, f)
        dst = os.path.join(backend_dir, f)
        print(f"Found: {src}")
        if f in ["localhost.crt", "localhost.key", "ca.crt"]:
            shutil.copy2(src, dst)
            print(f"Copied {f} to backend/")
else:
    print("Directory does not exist.")
