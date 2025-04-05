import hashlib
import uuid
import datetime
import json

def generate_license_key(user_info):
    """Generate a unique license key based on user info"""
    base = f"{user_info['name']}:{user_info['email']}:{user_info['hardware_id']}"
    return hashlib.sha256(base.encode()).hexdigest()[:16]

def create_license(user_info, duration_days=365):
    """Create a license file for the user"""
    license_key = generate_license_key(user_info)
    expiry_date = (datetime.datetime.now() + datetime.timedelta(days=duration_days)).strftime('%Y-%m-%d')
    
    license_data = {
        'license_key': license_key,
        'user_name': user_info['name'],
        'user_email': user_info['email'],
        'hardware_id': user_info['hardware_id'],
        'issue_date': datetime.datetime.now().strftime('%Y-%m-%d'),
        'expiry_date': expiry_date,
        'type': 'professional'
    }
    
    filename = f"license_{user_info['name'].lower().replace(' ', '_')}.json"
    with open(filename, 'w') as f:
        json.dump(license_data, f, indent=4)
    
    print(f"License generated successfully!")
    print(f"License Key: {license_key}")
    print(f"Expiry Date: {expiry_date}")

if __name__ == "__main__":
    user_info = {
        'name': input("Enter user name: "),
        'email': input("Enter user email: "),
        'hardware_id': input("Enter hardware ID from user's activation screen: ")
    }
    create_license(user_info)
