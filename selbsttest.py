"""Selbsttest: prueft das Regelwerk gegen die Testfaelle. Exit-Code 1 bei Abweichung."""
import json, sys
from klassifizierung import Classifier

spec = sys.argv[1] if len(sys.argv) > 1 else "klassifizierung_v1.json"
cases_file = sys.argv[2] if len(sys.argv) > 2 else "testfaelle.json"
c = Classifier(spec)
cases = json.load(open(cases_file, encoding="utf-8"))["faelle"]
bad = []
for t in cases:
    r = c.classify(t["account_no"], t["account_name"], t["account_group"])
    if r["category"] != t["erwartet"]:
        bad.append((t, r))
print(f"{len(cases) - len(bad)} von {len(cases)} Testfaellen bestanden")
for t, r in bad:
    print(f"  FEHLER {t['account_no']:10}{t['account_name'][:38]:40}"
          f"erwartet {t['erwartet']:26}erhalten {r['category']:26}({r['source']}, {r['rule_id']})")
sys.exit(1 if bad else 0)
