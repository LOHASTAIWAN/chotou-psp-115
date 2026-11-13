# -*- coding: utf-8 -*-
"""
GitHub Public Deployment & GitHub Pages Automated Publisher
依據使用者全域規範：自動建立 Public Repo、推送版本並啟用 GitHub Pages
"""
import ctypes
import ctypes.wintypes
import json
import os
import subprocess
import sys
import urllib.request
import urllib.error

REPO_NAME = "chotou-psp-115"
REPO_DESC = "115學年度 中投區高級中等學校實用技能學程 招生簡章與選填導引互動式 DocuSheet"

class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ('Flags', ctypes.wintypes.DWORD),
        ('Type', ctypes.wintypes.DWORD),
        ('TargetName', ctypes.wintypes.LPWSTR),
        ('Comment', ctypes.wintypes.LPWSTR),
        ('LastWritten', ctypes.wintypes.FILETIME),
        ('CredentialBlobSize', ctypes.wintypes.DWORD),
        ('CredentialBlob', ctypes.c_void_p),
        ('Persist', ctypes.wintypes.DWORD),
        ('AttributeCount', ctypes.wintypes.DWORD),
        ('Attributes', ctypes.c_void_p),
        ('TargetAlias', ctypes.wintypes.LPWSTR),
        ('UserName', ctypes.wintypes.LPWSTR),
    ]

def get_github_token():
    advapi32 = ctypes.windll.advapi32
    CredReadW = advapi32.CredReadW
    CredReadW.argtypes = [ctypes.wintypes.LPWSTR, ctypes.wintypes.DWORD, ctypes.wintypes.DWORD, ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
    CredReadW.restype = ctypes.wintypes.BOOL
    CredFree = advapi32.CredFree
    CredFree.argtypes = [ctypes.c_void_p]

    pcred = ctypes.POINTER(CREDENTIAL)()
    res = CredReadW('git:https://github.com', 1, 0, ctypes.byref(pcred))
    if not res:
        raise RuntimeError("Credential not found in Windows Credential Manager for git:https://github.com")
    
    blob = ctypes.string_at(pcred.contents.CredentialBlob, pcred.contents.CredentialBlobSize)
    token = blob.decode('utf-16-le').strip('\x00')
    CredFree(pcred)
    return token

def ensure_github_repo(token, username):
    url = f"https://api.github.com/repos/{username}/{REPO_NAME}"
    headers = {
        "Authorization": f"token {token}",
        "User-Agent": "Antigravity-Agent",
        "Accept": "application/vnd.github.v3+json"
    }

    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            print(f"[OK] Repository exists: {data['html_url']}")
            return data
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print(f"[INFO] Creating public repository '{REPO_NAME}' on GitHub...")
            create_url = "https://api.github.com/user/repos"
            payload = json.dumps({
                "name": REPO_NAME,
                "description": REPO_DESC,
                "private": False,
                "auto_init": False
            }).encode('utf-8')
            create_req = urllib.request.Request(create_url, data=payload, headers=headers, method="POST")
            with urllib.request.urlopen(create_req) as create_resp:
                new_repo = json.loads(create_resp.read().decode())
                print(f"[OK] Created public repository: {new_repo['html_url']}")
                return new_repo
        else:
            raise

def enable_github_pages(token, username):
    url = f"https://api.github.com/repos/{username}/{REPO_NAME}/pages"
    headers = {
        "Authorization": f"token {token}",
        "User-Agent": "Antigravity-Agent",
        "Accept": "application/vnd.github.v3+json"
    }
    
    # Check if already enabled
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            print(f"[OK] GitHub Pages is active: {data.get('html_url')}")
            return data.get('html_url')
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print("[INFO] Enabling GitHub Pages for main branch root...")
            payload = json.dumps({
                "source": {
                    "branch": "main",
                    "path": "/"
                }
            }).encode('utf-8')
            create_req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
            try:
                with urllib.request.urlopen(create_req) as create_resp:
                    res_data = json.loads(create_resp.read().decode())
                    pages_url = res_data.get('html_url', f"https://{username.lower()}.github.io/{REPO_NAME}/")
                    print(f"[OK] GitHub Pages enabled successfully: {pages_url}")
                    return pages_url
            except urllib.error.HTTPError as err:
                print(f"[WARN] Enable Pages response ({err.code}): {err.read().decode()}")
                return f"https://{username.lower()}.github.io/{REPO_NAME}/"
        else:
            print(f"[WARN] Pages check status ({e.code}): {e.read().decode()}")
            return f"https://{username.lower()}.github.io/{REPO_NAME}/"

def main():
    token = get_github_token()
    username = "LOHASTAIWAN"
    
    # 1. 確保遠端倉庫存在
    repo_data = ensure_github_repo(token, username)
    repo_url = repo_data['html_url']

    # 2. 本地 Git 初始化或設定
    is_git = os.path.exists(".git")
    if not is_git:
        print("[INFO] Initializing git repository on main branch...")
        subprocess.run(["git", "init", "-b", "main"], check=True)
        subprocess.run(["git", "config", "user.name", "LOHASTAIWAN"], check=True)
        subprocess.run(["git", "config", "user.email", "lohas007@gmail.com"], check=True)

    # 3. Commit
    commit_msg = sys.argv[1] if len(sys.argv) > 1 else "feat: 115學年度中投區實用技能學程招生簡章DocuSheet完整彙編與水電科專題發布"
    subprocess.run(["git", "add", "."], check=True)
    
    # Commit check
    status = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
    if status.stdout.strip():
        subprocess.run(["git", "commit", "-m", commit_msg], check=True)
        print(f"[OK] Git commit created: {commit_msg}")
    else:
        print("[INFO] Working tree clean, nothing to commit.")

    # 4. Push
    authed_remote = f"https://{username}:{token}@github.com/{username}/{REPO_NAME}.git"
    clean_remote = f"https://github.com/{username}/{REPO_NAME}.git"

    # Set remote
    remotes = subprocess.run(["git", "remote"], capture_output=True, text=True).stdout.split()
    if "origin" in remotes:
        subprocess.run(["git", "remote", "set-url", "origin", authed_remote], check=True)
    else:
        subprocess.run(["git", "remote", "add", "origin", authed_remote], check=True)

    print(f"[INFO] Pushing commits to GitHub ({REPO_NAME}:main)...")
    push_res = subprocess.run(["git", "push", "-u", "origin", "main"], capture_output=True, text=True)
    
    # 復原 remote-url 避免暴露 token
    subprocess.run(["git", "remote", "set-url", "origin", clean_remote], check=True)

    if push_res.returncode != 0:
        print("[ERROR] Push failed:")
        print(push_res.stderr)
        sys.exit(push_res.returncode)
    else:
        print("[OK] Push succeeded!")

    # 5. 啟用 GitHub Pages
    pages_url = enable_github_pages(token, username)
    print("\n" + "="*60)
    print("🚀 部署成功！公開存取資訊：")
    print(f"🔗 GitHub 倉庫：{repo_url}")
    print(f"🌐 GitHub Pages 公開網址：{pages_url}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
