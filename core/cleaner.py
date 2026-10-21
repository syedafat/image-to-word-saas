import os
import time

def cleanup_old_files(folder_path: str, max_age_seconds: int = 3600):
    """
    निर्दिष्ट फोल्डर में से निश्चित समय (default: 1 घंटा) से पुरानी फाइलों को खुद-ब-खुद हटाता है।
    """
    now = time.time()
    
    if not os.path.exists(folder_path):
        return

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # केवल फाइलों को प्रोसेस करें (सब-फोल्डर्स को छोड़कर)
        if os.path.isfile(file_path):
            file_age = now - os.path.getmtime(file_path)
            
            # यदि फाइल तय समय से पुरानी है तो डिलीट करें
            if file_age > max_age_seconds:
                try:
                    os.remove(file_path)
                    print(f"[Cleanup] Deleted old file: {file_path}")
                except Exception as e:
                    print(f"[Cleanup Error] Could not delete {file_path}: {e}")