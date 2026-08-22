import subprocess
import os
import re

def rearm_office():
    office_path = r"C:\Program Files\Microsoft Office\Office16"
    
    if not os.path.exists(office_path):
        print(f"Office not found at {office_path}")
        return
    
    os.chdir(office_path)
    
    # Check status
    print("Checking Office status...")
    result = subprocess.run('cscript ospp.vbs /dstatus', shell=True, capture_output=True, text=True)
    print(result.stdout)
    
    # Extract SKU ID
    sku_match = re.search(r'SKU ID:\s*([a-f0-9-]+)', result.stdout, re.IGNORECASE)
    
    print("\n" + "="*50)
    
    if sku_match:
        sku_id = sku_match.group(1)
        print(f"🔑 SKU ID Found: {sku_id}")
        print(f"📋 Command: ospprearm {sku_id}")
        print("\n" + "="*50)
        
        # Rearm with SKU ID
        print(f"Rearming Office with SKU: {sku_id}...")
        subprocess.run(f'ospprearm {sku_id}', shell=True)
    else:
        print("⚠️  No SKU ID found, rearming all...")
        subprocess.run('ospprearm', shell=True)
    
    print("\n" + "="*50)
    print("Office rearm completed!")

if __name__ == "__main__":
    rearm_office()
    input("Press Enter to exit...")
