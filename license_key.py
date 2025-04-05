# Keep this script private - do not distribute
import hashlib
import uuid
import platform

def generate_license_key(hardware_id):
    return hashlib.sha256(hardware_id.encode()).hexdigest()[:16]

# Get hardware ID from user's machine
hardware_id = input("Enter user's Hardware ID: ")
license_key = generate_license_key(hardware_id)
print(f"License Key: {license_key}")