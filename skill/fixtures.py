#!/usr/bin/env python3
"""
Realistic mock M365 documents for testing the m365-surface-commander skill across surfaces.

Each fixture is a small "bridge": it renders a <doc_state> snapshot the way the add-in would, and
applies the skill's commands (read/search/set/suggest/comment/...) returning the text the add-in
would put in a ```result block. Three surfaces:

  - ExcelSales   : a quarterly sales-analysis sheet with a planted anomaly (data analysis)
  - OutlookThread: an email thread asking to confirm a meeting + send an agenda (email context)
  - ContractDoc  : an MSA / CLM-style contract with risky clauses (contract review)
"""
from __future__ import annotations

import re


class ExcelSales:
    surface = "excel"
    caps = "outline, read, search, set, format, comment, done"

    def __init__(self):
        # Region, Month, Revenue, Cost  (rows 2..9). Row 8 (West Mar) is an anomalous spike.
        self.headers = ["Region", "Month", "Revenue", "Cost"]
        self.rows = [
            ["East", "Jan", 100, 60], ["West", "Jan", 80, 50],
            ["East", "Feb", 120, 70], ["West", "Feb", 90, 55],
            ["East", "Mar", 140, 80], ["North", "Mar", 70, 40],
            ["West", "Mar", 900, 60],  # <- anomaly: revenue spike
            ["North", "Feb", 60, 35],
        ]
        self.cells = {}
        cols = "ABCD"
        for j, h in enumerate(self.headers):
            self.cells[f"{cols[j]}1"] = h
        for i, row in enumerate(self.rows, start=2):
            for j, val in enumerate(row):
                self.cells[f"{cols[j]}{i}"] = str(val)
        self.nrows = len(self.rows) + 1
        self.writes, self.comments = [], []

    def doc_state(self):
        out = [f'Excel — sheet "Sales", rows 1-{self.nrows}, columns A-D:']
        for i in range(1, self.nrows + 1):
            cs = [f"{c}{i}={self.cells[f'{c}{i}']}" for c in "ABCD" if f"{c}{i}" in self.cells]
            out.append("  " + "  ".join(cs))
        out.append("  F2=<empty>  F3=<empty>  F4=<empty>  (summary area)")
        return "<doc_state>\n" + "\n".join(out) + "\n</doc_state>"

    def _read(self, sel):
        sel = sel.split("!")[-1]
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", sel)
        if not m:
            return f"{sel}={self.cells.get(sel,'<empty>')}"
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        got = [f"{chr(ci)}{r}={self.cells[f'{chr(ci)}{r}']}"
               for r in range(r1, r2 + 1) for ci in range(ord(c1), ord(c2) + 1)
               if f"{chr(ci)}{r}" in self.cells]
        return " ".join(got) or f"(empty: {sel})"

    def apply(self, cmd):
        v = cmd["verb"]
        if v == "outline":
            return "Sales: A=Region B=Month C=Revenue D=Cost; rows 2-9; summary area F2:F4"
        if v == "read":
            return self._read(cmd["selector"])
        if v == "search":
            t = cmd["text"].lower()
            return "matches: " + (", ".join(k for k, val in self.cells.items()
                                             if t in str(val).lower()) or "(none)")
        if v == "set":
            cell = cmd["cell"].split("!")[-1]
            self.cells[cell] = cmd["value"]
            self.writes.append((cell, cmd["value"]))
            return f"OK set {cell} = {cmd['value']}"
        if v == "format":
            return f"OK formatted {cmd['range']}"
        if v == "comment":
            self.comments.append((cmd["selector"], cmd["text"]))
            return f"OK comment on {cmd['selector']}"
        return f"OK {v}"


class OutlookThread:
    surface = "outlook"
    caps = "read, search, mail, compose, done"

    def __init__(self):
        self.subject = "Re: Q3 planning sync"
        self.frm = "priya@contoso.com"
        self.body = [
            "Hi — can we lock the Q3 planning sync?",
            "I'm open Thursday 3pm or Friday 10am AEST.",
            "Also, do you have the latest revenue deck to pre-read?",
        ]
        self.drafts = []

    def doc_state(self):
        b = "\n  ".join(self.body)
        return (f"<doc_state>\nOutlook — open message:\n  Subject: {self.subject}\n"
                f"  From: {self.frm}\n  Body:\n  {b}\n</doc_state>")

    def apply(self, cmd):
        v = cmd["verb"]
        if v == "read":
            return f"Subject: {self.subject} | From: {self.frm} | Body: {' '.join(self.body)}"
        if v == "search":
            t = cmd["text"].lower()
            hits = [ln for ln in self.body if t in ln.lower()]
            return "matches: " + (" | ".join(hits) or "(none)")
        if v == "mail":
            self.drafts.append(("reply", cmd["body"]))
            return "OK staged reply draft (not sent)"
        if v == "compose":
            self.drafts.append(("new", cmd.get("subject", ""), cmd["body"]))
            return "OK opened new draft (unaddressed, not sent)"
        return f"OK {v}"


class ContractDoc:
    surface = "word"
    caps = "outline, read, search, suggest, comment, done"

    def __init__(self):
        # CLM / MSA style clauses; a couple are deliberately risky.
        self.clauses = {
            "1. Term": "This Agreement commences on the Effective Date and continues for twelve (12) months.",
            "2. Fees": "Customer shall pay all fees within ninety (90) days of invoice.",  # risky: 90 days
            "3. Liability": "In no event shall either party's liability exceed the fees paid in the prior 12 months.",
            "4. Auto-Renewal": "This Agreement automatically renews for successive 12-month terms unless cancelled.",  # risky: no notice window
            "5. Governing Law": "This Agreement is governed by the laws of the State of Delaware.",
        }
        self.suggestions, self.comments = [], []

    def doc_state(self):
        out = ["Word — contract (MSA), section outline + text:"]
        for h, t in self.clauses.items():
            out.append(f"  ## {h}\n  {t}")
        return "<doc_state>\n" + "\n".join(out) + "\n</doc_state>"

    def apply(self, cmd):
        v = cmd["verb"]
        if v == "outline":
            return "Sections: " + "; ".join(self.clauses.keys())
        if v == "read":
            sel = cmd["selector"].strip().strip('"')
            if not sel:
                return " ".join(self.clauses.values())
            for h, t in self.clauses.items():
                if sel.lower() in h.lower() or sel.lower() in t.lower():
                    return f"{h}: {t}"
            return "(no matching section)"
        if v == "search":
            t = cmd["text"].lower()
            hits = [h for h, txt in self.clauses.items() if t in txt.lower() or t in h.lower()]
            return "matches: " + (", ".join(hits) or "(none)")
        if v == "suggest":
            old = cmd["oldText"]
            found = any(old in t for t in self.clauses.values())
            self.suggestions.append((cmd["oldText"], cmd["newText"]))
            return ("OK tracked change staged" if found
                    else f"error: anchor text not found verbatim: \"{old[:40]}\" — re-read and retry")
        if v == "comment":
            self.comments.append((cmd["selector"], cmd["text"]))
            return f"OK comment anchored on \"{cmd['selector'][:40]}\""
        return f"OK {v}"


FIXTURES = {"excel": ExcelSales, "email": OutlookThread, "contract": ContractDoc}

TASKS = {
    "excel": ("Add each region's total Revenue to the summary area (F2 East, F3 West, F4 North) "
              "and add a comment on the single anomalous revenue row."),
    "email": ("Reply to confirm Thursday 3pm AEST and ask them to send the revenue deck as a "
              "pre-read."),
    "contract": ("Review the contract. Propose a tracked change tightening the payment term to "
                 "30 days, and comment on any clause that is risky for the customer."),
}


if __name__ == "__main__":
    for key, cls in FIXTURES.items():
        f = cls()
        print(f"\n===== {key}  (surface={f.surface}) =====")
        print(f.doc_state())
        print("TASK:", TASKS[key])
