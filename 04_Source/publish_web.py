# -*- coding: utf-8 -*-
r"""
PUBLISH_WEB - rebuild the site in 06_Web and push it to GitHub Pages
(the public URL comes from github_owner/github_repo in 03_System/config.json).

Usage: python publish_web.py [plan_month YYYY-MM]   (default: current month)
Steps: build_web.py (which FIRST regenerates every downloadable report via
analyze_sms.py + plan_month.py - hard-fails if either errors) -> git add/commit
-> git push -> wait Pages build.
ASCII-only console output.
"""
import os, sys, subprocess, time, datetime as dt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(BASE, "06_Web")
from mfconfig import repo_slug, site_url  # noqa: E402

MONTH = sys.argv[1] if len(sys.argv) > 1 else dt.date.today().strftime("%Y-%m")
REPO = repo_slug()
SITE = site_url()


def run(cmd, cwd=WEB, ok_codes=(0,)):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def main():
    # 1. rebuild site
    rc, out = run([sys.executable, os.path.join(BASE, "04_Source", "build_web.py"), MONTH])
    print(out.strip())
    if rc != 0:
        print("BUILD FAILED")
        sys.exit(1)

    # 2. commit
    run(["git", "add", "-A"])
    msg = "update " + dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " (plan " + MONTH + ")"
    rc, out = run(["git", "commit", "-m", msg])
    if "nothing to commit" in out:
        print("no changes to publish")
    elif rc != 0:
        print("COMMIT FAILED:\n" + out)
        sys.exit(1)
    else:
        print("committed: " + msg)

    # 3. push
    rc, out = run(["git", "push", "origin", "main"])
    if rc != 0:
        print("PUSH FAILED:\n" + out)
        sys.exit(1)
    print("pushed to origin/main")

    # 4. wait for Pages build
    if not REPO:
        print("NOTE: github_owner/github_repo not set in 03_System/config.json")
        print("      - cannot check the Pages build status; the push above is done.")
        print("      - set them and enable Pages (branch: main, folder: /docs or /)")
        return
    print("waiting for GitHub Pages build...")
    for _ in range(30):
        rc, out = run(["gh", "api", "repos/" + REPO + "/pages", "--jq", ".status"])
        status = out.strip().splitlines()[-1] if out.strip() else "?"
        if status == "built":
            print("STATUS: built")
            print("LIVE: " + SITE)
            return
        time.sleep(6)
    print("STATUS: still " + status + " - check later: " + SITE)


if __name__ == "__main__":
    main()
