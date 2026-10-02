# -*- coding: utf-8 -*-
"""
mfconfig - shared config loader for every MyFinance template script.

All personal/system settings live in one file: 03_System/config.json
(created by init_project.py). Scripts import this module instead of
hardcoding anything.

Keys:
  name        display name of the owner (site title)
  passcode    site gate passcode (never stored plain in the site)
  income      monthly income in local currency units
  currency    currency label, e.g. "ج.م"
  city        city used in the market section title
  year        report year used in output file names
  github_owner / github_repo   e.g. "myuser" / "myfinance"
  weekly_cap_fallback  weekly cash cap if no plan file exists (0 = derive)
"""
import datetime as dt
import hashlib
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CFG_PATH = os.path.join(BASE, "03_System", "config.json")

_DEFAULTS = {
    "name": "MyFinance",
    "passcode": "change-me",
    "income": 0,
    "currency": "EGP",
    "city": "Cairo",
    "github_owner": "",
    "github_repo": "",
    "weekly_cap_fallback": 0,
}


def load():
    cfg = dict(_DEFAULTS)
    cfg["year"] = dt.date.today().year
    if os.path.isfile(CFG_PATH):
        try:
            with open(CFG_PATH, "r", encoding="utf-8") as fh:
                cfg.update(json.load(fh))
        except Exception as exc:  # pragma: no cover - defensive
            print("WARN: cannot read config.json ({}) - using defaults".format(exc))
    return cfg


CFG = load()
YEAR = int(CFG.get("year") or dt.date.today().year)
GATE = hashlib.sha256(str(CFG["passcode"]).encode("utf-8")).hexdigest()
TOKEN_SALT = str(CFG["passcode"]) + ":"


def site_url():
    owner = str(CFG.get("github_owner") or "").strip()
    repo = str(CFG.get("github_repo") or "").strip()
    if owner and repo:
        return "https://{}.github.io/{}/".format(owner, repo)
    return ""


def repo_slug():
    owner = str(CFG.get("github_owner") or "").strip()
    repo = str(CFG.get("github_repo") or "").strip()
    if owner and repo:
        return "{}/{}".format(owner, repo)
    return ""
