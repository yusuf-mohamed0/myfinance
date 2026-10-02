# -*- coding: utf-8 -*-
r"""
INIT_PROJECT - one-time setup for a fresh copy of the MyFinance system.
Runs on macOS (python3) and Windows (python). Fully offline except the
optional GitHub steps.

What it does:
  1. Collects your settings (flags, or interactive prompts when run in a
     terminal without --yes; never blocks on input() in flags/batch mode)
  2. Writes 03_System/config.json  (name, passcode, income, city, currency,
     github_owner, github_repo, year)
  3. Generates a sample SMS file 01_Data/sms_raw_sample.txt (delete it once
     you paste your real bank SMS) so the whole pipeline can run end-to-end
  4. Runs the pipeline:
       build_excel.py -> analyze_sms.py -> plan_month.py -> build_web.py
  5. Optional (--no-repo to skip): creates YOUR GitHub repo, pushes 06_Web
     and enables GitHub Pages
  6. Runs verify_system.py (green checklist)

Usage examples:
  python3 init_project.py --yes --name Ibrahim --passcode my-secret \
      --income 15000 --no-repo
  python3 init_project.py                      (interactive)

System lock policy: this script never edits files under 04_Source, and
verify_system.py fails if anyone changes them (03_System/source_manifest.json).
Changing the system requires the system owner's permission.
"""
import argparse
import datetime as dt
import glob
import json
import os
import shutil
import subprocess
import sys

# Arabic output must not crash a cp1252 Windows console
def _mf_utf8():
    for _n in ("stdout", "stderr"):
        _st = getattr(sys, _n, None)
        if _st is not None:
            try:
                _st.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
_mf_utf8()

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "04_Source")
DATA = os.path.join(BASE, "01_Data")
REP = os.path.join(BASE, "02_Reports")
SYS = os.path.join(BASE, "03_System")
WEB = os.path.join(BASE, "06_Web")


# ---------------------------------------------------------------- helpers
def ask(label, default=""):
    """Prompt once; safe on EOF/Ctrl-C (returns default)."""
    try:
        suffix = " [{}] ".format(default) if default != "" else ": "
        v = input(label + suffix)
    except (EOFError, KeyboardInterrupt):
        print("")
        return default
    v = v.strip()
    return v if v else default


def yn(label, default=True):
    v = ask(label + " (y/n)", "y" if default else "n").lower()
    return v.startswith("y")


def gh_login():
    """Logged-in GitHub login, or '' when gh is missing/not authed."""
    try:
        p = subprocess.run(["gh", "api", "user", "--jq", ".login"],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=30)
        if p.returncode == 0 and p.stdout.strip():
            return p.stdout.strip()
    except Exception:
        pass
    return ""


def run_step(title, script, *args):
    print("")
    print("=== " + title)
    cmd = [sys.executable, os.path.join(SRC, script)] + [str(a) for a in args]
    p = subprocess.run(cmd, cwd=BASE)
    if p.returncode != 0:
        print("STEP FAILED: " + script + " (exit {})".format(p.returncode))
    return p.returncode == 0


def write_config(cfg):
    os.makedirs(SYS, exist_ok=True)
    path = os.path.join(SYS, "config.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
    print("wrote 03_System/config.json")


# ---------------------------------------------------------------- harden ACL
def harden():
    """Create the remaining project folders, then lock the tree down to the
    current Windows account only (NTFS). Best-effort: needs Windows + icacls."""
    for d in ("01_Data", "02_Reports", "03_System", "04_Source", "05_Docs",
              "06_Web"):
        os.makedirs(os.path.join(BASE, d), exist_ok=True)

    if not sys.platform.startswith("win"):
        print("NOTE: NTFS lockdown skipped (not Windows) - the tree is still "
              "entirely local.")
        return False

    owner = os.environ.get("USERNAME") or os.environ.get("USER") or ""
    if not owner:
        print("NOTE: could not detect your Windows account - ACL not applied.")
        return False

    acct = "TL\\" + owner if not owner.startswith("TL\\") else owner
    icacls = shutil.which("icacls")
    if not icacls:
        print("NOTE: icacls not found - ACL not applied.")
        return False

    targets = ["01_Data", "02_Reports", "03_System", "05_Docs", "06_Web"]
    for t in targets:
        p = os.path.join(BASE, t)
        try:
            subprocess.run(
                [icacls, p, "/inheritance:r", "/grant:r", "{}:(OI)(CI)F".format(acct),
                 "/remove", "*S-1-1-0", "*S-1-5-11", "*S-1-5-32-545"],
                capture_output=True, text=True, errors="replace")
            print("locked {} ({} only)".format(t, acct))
        except Exception as exc:
            print("NOTE: could not lock {} ({})".format(t, exc))
    return True


# ---------------------------------------------------------------- sample SMS
def write_sample_sms():
    """NBE-style sample that matches analyze_sms.py regexes exactly and
    balances the running-balance chain (drift = 0). Dynamic dates."""
    os.makedirs(DATA, exist_ok=True)
    path = os.path.join(DATA, "sms_raw_sample.txt")
    if glob.glob(os.path.join(DATA, "sms_raw*.txt")):
        print("SMS input already present - sample not written")
        return False

    today = dt.date.today()

    def day(offset):
        x = today - dt.timedelta(days=offset)
        return "{}/{}/{}".format(x.day, x.month, x.year)

    lines = [
        # earliest first; first debit carries the balance anchor
        "تم خصم 250.00 EGP عند ATM - MAADI يوم {} المتاح 4750.00".format(day(15)),
        "تم خصم 180.00 EGP عند OLA ENERGY يوم {} المتاح 4570.00".format(day(14)),
        "تم إضافة تحويل وارد مبلغ 10000.00 من شركة النيل رقم مرجعي 123456 يوم {}".format(day(13)),
        "تم خصم 300.00 EGP عند TALABAT يوم {} المتاح 14270.00".format(day(12)),
        "تم رد مبلغ 40.00 EGP من صيدلة تجريبية يوم {}".format(day(11)),
        "تم خصم 600.00 EGP عند VODAFONE يوم {} المتاح 13710.00".format(day(10)),
        "ناسف لعدم إتمام المعاملة بمبلغ 300.00 EGP يوم {}".format(day(10)),
        "رمز التحقق الخاص بك 458912 صالح لمدة 5 دقائق",
        "Your account statement is ready in the mobile app.",
    ]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("wrote 01_Data/sms_raw_sample.txt ({} lines) - SAMPLE, "
          "delete it when you paste your real SMS".format(len(lines)))
    return True


# ---------------------------------------------------------------- GitHub
def setup_repo(owner, repo):
    """git init 06_Web, create repo, push, enable Pages. Returns site URL."""
    slug = "{}/{}".format(owner, repo)
    site = "https://{}.github.io/{}/".format(owner, repo)

    if not os.path.isdir(WEB):
        print("no 06_Web - run the build first")
        return ""

    def git(*args):
        return subprocess.run(["git"] + list(args), cwd=WEB,
                              capture_output=True, text=True,
                              encoding="utf-8", errors="replace")

    if not os.path.isdir(os.path.join(WEB, ".git")):
        r = git("init", "-b", "main")
        if r.returncode != 0:  # older git without -b
            git("init")
            git("checkout", "-b", "main")
    git("add", "-A")
    git("commit", "-m", "init MyFinance site " +
        dt.datetime.now().strftime("%Y-%m-%d %H:%M"))

    # create + push (or just push if the repo already exists)
    r = subprocess.run(["gh", "repo", "create", slug, "--public",
                        "--source", ".", "--push"],
                       cwd=WEB, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        git("remote", "add", "origin", "https://github.com/{}.git".format(slug))
        p = git("push", "-u", "origin", "main")
        if p.returncode != 0:
            print("PUSH FAILED:\n" + (p.stdout or "") + (p.stderr or ""))
            print("Fix: cd 06_Web && git push -u origin main")
            return ""
        print("pushed to " + slug + " (remote already existed)")
    else:
        print("created + pushed " + slug)

    # enable Pages (idempotent)
    try:
        get = subprocess.run(["gh", "api", "repos/" + slug + "/pages"],
                             capture_output=True, text=True, encoding="utf-8",
                             errors="replace")
        if get.returncode != 0:
            subprocess.run(["gh", "api", "-X", "POST",
                            "repos/" + slug + "/pages",
                            "-f", "source[branch]=main",
                            "-f", "source[path]=/"],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace")
        # wait briefly for the first build
        for _ in range(20):
            st = subprocess.run(["gh", "api", "repos/" + slug + "/pages",
                                 "--jq", ".status"],
                                capture_output=True, text=True,
                                encoding="utf-8", errors="replace")
            if st.stdout.strip() == "built":
                break
            dt.timedelta(seconds=0)
            import time
            time.sleep(4)
    except Exception as exc:
        print("pages api note: {}".format(exc))

    print("SITE: " + site)
    return site


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(
        description="One-time MyFinance setup (config, sample data, build, GitHub).")
    ap.add_argument("--yes", action="store_true",
                    help="no prompts; use flags/defaults (never blocks)")
    ap.add_argument("--name", default="", help="display name / site title")
    ap.add_argument("--passcode", default="", help="site gate passcode")
    ap.add_argument("--income", type=float, default=None,
                    help="monthly income in local currency")
    ap.add_argument("--city", default="", help="city for the market section")
    ap.add_argument("--currency", default="", help="currency label (e.g. 'EGP')")
    ap.add_argument("--owner", default="", help="GitHub username")
    ap.add_argument("--repo", default="myfinance", help="GitHub repo name to create")
    ap.add_argument("--no-repo", action="store_true",
                    help="skip all GitHub steps (everything stays local)")
    ap.add_argument("--skip-build", action="store_true",
                    help="only write config + sample SMS, run no builds")
    args = ap.parse_args()

    interactive = (not args.yes) and sys.stdin.isatty()
    detected = gh_login()

    # ---- 1. settings ----
    cfg = {
        "name": args.name,
        "passcode": args.passcode,
        "income": args.income,
        "city": args.city,
        "currency": args.currency,
        "github_owner": args.owner,
        "github_repo": "" if args.no_repo else args.repo,
        "year": dt.date.today().year,
        "weekly_cap_fallback": 0,
    }

    if interactive:
        print("MyFinance setup - press Enter to accept the [default] values.")
        if not cfg["name"]:
            cfg["name"] = ask("Display name / site title", detected or "MyFinance")
        if not cfg["passcode"]:
            cfg["passcode"] = ask("Site passcode (you type it to unlock the site)",
                                  "change-me")
        if cfg["income"] is None:
            v = ask("Monthly income", "0")
            try:
                cfg["income"] = float(v.replace(",", ""))
            except ValueError:
                cfg["income"] = 0.0
        if not cfg["city"]:
            cfg["city"] = ask("City (market section)", "Cairo")
        if not cfg["currency"]:
            cfg["currency"] = ask("Currency label", "EGP")
        if not args.no_repo:
            # default the owner to the detected gh login ONLY if the user agrees;
            # otherwise empty (local-only) so we never assume someone's account
            if not cfg["github_owner"]:
                cfg["github_owner"] = ask("GitHub username (blank = local only)",
                                         detected or "")
            repo_default = args.repo or "myfinance"
            cfg["github_repo"] = ask("GitHub repo name (blank = skip GitHub)",
                                     repo_default)
            if not cfg["github_repo"]:
                cfg["github_owner"] = ""
            else:
                create = yn("Create/push the repo + enable Pages now", True)
                if not create:
                    cfg["github_repo"] = ""
                    cfg["github_owner"] = ""
    else:
        # batch mode: fill sensible defaults, never prompt
        if not cfg["name"]:
            cfg["name"] = detected or "MyFinance"
        if not cfg["passcode"]:
            cfg["passcode"] = "change-me"
        if cfg["income"] is None:
            cfg["income"] = 0.0
        if not cfg["city"]:
            cfg["city"] = "Cairo"
        if not cfg["currency"]:
            cfg["currency"] = "EGP"
        if not cfg["github_owner"]:
            cfg["github_owner"] = detected or ""
        if args.no_repo:
            cfg["github_owner"] = ""
        if not cfg["github_owner"]:
            # cannot push without an owner - stay local, tell the user later
            cfg["github_repo"] = ""
        elif not cfg["github_repo"]:
            # owner known but no repo requested - stay local by default
            cfg["github_repo"] = ""
        if cfg["passcode"] == "change-me":
            print("WARN: no --passcode given, using placeholder 'change-me' - "
                  "set a real one before publishing.")

    write_config(cfg)

    print("\n=== lockdown (NTFS: your Windows account only)")
    harden()

    if cfg["passcode"] in ("", "change-me") and not args.no_repo:
        print("HOLD: refusing to publish with a placeholder passcode.")
        print("      re-run with --passcode <your-secret> (site would be open).")
        cfg["github_repo"] = ""

    # ---- 2. sample data ----
    write_sample_sms()

    if args.skip_build:
        print("\nConfig done (--skip-build). Next: python3 init_project.py")
        return 0

    # ---- 3. pipeline ----
    ok = True
    ok &= run_step("1/4 build Excel workbook", "build_excel.py")
    sms_present = bool(glob.glob(os.path.join(DATA, "sms_raw*.txt")))
    if sms_present:
        ok &= run_step("2/4 analyze SMS -> reports", "analyze_sms.py")
    else:
        print("\n=== 2/4 analyze SMS - SKIPPED (no 01_Data/sms_raw*.txt)")
    summary = os.path.join(REP, "summary_{}.json".format(cfg["year"]))
    if os.path.isfile(summary) and cfg["income"]:
        month = dt.date.today().strftime("%Y-%m")
        ok &= run_step("3/4 monthly plan " + month, "plan_month.py",
                       cfg["income"], month)
    else:
        print("\n=== 3/4 monthly plan - SKIPPED (need SMS summary + income)")
    if os.path.isfile(summary):
        ok &= run_step("4/4 build website", "build_web.py")
    else:
        print("\n=== 4/4 build website - SKIPPED (run analyze_sms first)")
        ok = False

    # ---- 4. GitHub ----
    site = ""
    if cfg["github_owner"] and cfg["github_repo"]:
        print("\n=== GitHub: push site + enable Pages")
        site = setup_repo(cfg["github_owner"], cfg["github_repo"])
        if not site:
            print("You can do it later:")
            print("  1. gh auth login          (your GitHub account)")
            print("  2. fill github_owner / github_repo in 03_System/config.json")
            print("  3. python3 04_Source/publish_web.py")
    elif not args.no_repo:
        print("\nGitHub skipped (no owner detected). To publish later:")
        print("  gh auth login  ->  re-run init_project.py  or set config.json")

    # ---- 5. verify ----
    print("")
    verified = run_step("verify", "verify_system.py")

    # ---- 6. wrap up ----
    print("")
    print("-" * 60)
    print("SETUP " + ("COMPLETE" if (ok and verified) else "FINISHED WITH WARNINGS"))
    if site:
        print("site     : " + site)
        print("passcode : " + ("(the one you set)" if cfg["passcode"] != "change-me"
                               else "change-me - CHANGE IT"))
    print("")
    print("NEXT STEPS")
    print("  1. Delete 01_Data/sms_raw_sample.txt and paste your own bank SMS")
    print("     as 01_Data/sms_raw_<something>.txt  (newest file wins).")
    print("  2. Re-run: analyze_sms -> plan_month <income> -> build_web -> publish_web")
    print("  3. Every Saturday: update 02_Reports/market_watch.json (prices),")
    print("     then build_web + publish_web.")
    print("  4. Do NOT edit 04_Source - verify_system.py fails if files change.")
    print("  5. AI assistant: install OpenCode (https://opencode.ai) and run")
    print("     'opencode' in this folder - it loads AGENTS.md + the myfinance")
    print("     skill, and system files are read-only for the agent by rules.")
    return 0 if (ok and verified) else 1


if __name__ == "__main__":
    sys.exit(main())
