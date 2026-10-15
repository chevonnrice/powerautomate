"""Static checks for the flow definitions in ../sourcecode.

Catches the mistakes that otherwise only show up on import or on the first run:
broken runAfter chains, references to actions/variables/parameters that do not exist,
connection references missing from Customizations.xml, and unbalanced expressions.

    python scripts/validate.py
"""
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "sourcecode")

errors = []


def err(flow, msg):
    errors.append(f"{flow}: {msg}")


def containers(actions, path="root"):
    """Yield (path, actions-dict) for every action container in the definition."""
    yield path, actions
    for name, a in actions.items():
        if "actions" in a:
            yield from containers(a["actions"], f"{path}/{name}")
        if "else" in a:
            yield from containers(a["else"].get("actions", {}), f"{path}/{name}/else")
        for cname, case in a.get("cases", {}).items():
            yield from containers(case.get("actions", {}), f"{path}/{name}/{cname}")
        if "default" in a:
            yield from containers(a["default"].get("actions", {}), f"{path}/{name}/default")


def strings(node):
    if isinstance(node, str):
        yield node
    elif isinstance(node, dict):
        for k, v in node.items():
            yield k
            yield from strings(v)
    elif isinstance(node, list):
        for v in node:
            yield from strings(v)


def expressions(s):
    """Return the expression bodies inside a workflow string."""
    if s.startswith("@") and not s.startswith("@{"):
        return [s[1:]]
    out, i = [], 0
    while True:
        i = s.find("@{", i)
        if i < 0:
            return out
        depth, j, quoted = 0, i + 1, False
        while j < len(s):
            c = s[j]
            if c == "'":
                quoted = not quoted
            elif not quoted and c == "{":
                depth += 1
            elif not quoted and c == "}":
                depth -= 1
                if depth == 0:
                    break
            j += 1
        out.append(s[i + 2 : j])
        i = j


def balanced(expr):
    depth, quoted = 0, False
    for c in expr:
        if c == "'":
            quoted = not quoted
        elif not quoted:
            if c in "([":
                depth += 1
            elif c in ")]":
                depth -= 1
                if depth < 0:
                    return False
    return depth == 0 and not quoted


with open(os.path.join(SRC, "Other", "Customizations.xml"), encoding="utf-8-sig") as f:
    customizations = f.read()
declared_refs = set(re.findall(r'connectionreferencelogicalname="([^"]+)"', customizations))
env_defs = {os.path.basename(d) for d in glob.glob(os.path.join(SRC, "environmentvariabledefinitions", "*"))}
with open(os.path.join(SRC, "Other", "Solution.xml"), encoding="utf-8-sig") as f:
    solution = f.read()

flows = sorted(glob.glob(os.path.join(SRC, "Workflows", "*.json")))
if not flows:
    errors.append("no workflows found")

for path in flows:
    flow = os.path.basename(path)
    with open(path, encoding="utf-8") as f:
        wf = json.load(f)
    props = wf["properties"]
    definition = props["definition"]
    conn_refs = props["connectionReferences"]

    gid = re.search(r"-([0-9A-F-]{36})\.json$", flow).group(1).lower()
    if f"{{{gid}}}" not in solution:
        err(flow, "workflow id is not a RootComponent in Solution.xml")
    if not os.path.exists(path + ".data.xml"):
        err(flow, "missing .data.xml")

    for api, ref in conn_refs.items():
        if ref["connection"]["connectionReferenceLogicalName"] not in declared_refs:
            err(flow, f"connection reference {ref['connection']['connectionReferenceLogicalName']} not in Customizations.xml")

    params = definition["parameters"]
    for name, prm in params.items():
        schema = prm.get("metadata", {}).get("schemaName")
        if schema and schema not in env_defs:
            err(flow, f"parameter {name} has no environment variable definition {schema}")

    all_actions = set(definition["triggers"])
    variables = set()
    for cpath, actions in containers(definition["actions"]):
        for name, a in actions.items():
            if name in all_actions:
                err(flow, f"duplicate action name {name}")
            all_actions.add(name)
            if a["type"] == "InitializeVariable":
                if cpath != "root":
                    err(flow, f"{name}: variables must be initialized at the top level")
                variables.update(v["name"] for v in a["inputs"]["variables"])
            for dep in a.get("runAfter", {}):
                if dep not in actions:
                    err(flow, f"{cpath}/{name}: runAfter '{dep}' is not in the same container")
            host = a.get("inputs", {}).get("host", {}) if isinstance(a.get("inputs"), dict) else {}
            if host and host["connectionName"] not in conn_refs:
                err(flow, f"{name}: connection {host['connectionName']} not in connectionReferences")
            if a["type"] == "SetVariable":
                own = a["inputs"]["name"]
                if f"variables('{own}')" in json.dumps(a["inputs"]["value"]):
                    err(flow, f"{name}: a variable cannot reference itself in Set variable")

    for trig in definition["triggers"].values():
        host = trig.get("inputs", {}).get("host")
        if host and host["connectionName"] not in conn_refs:
            err(flow, f"trigger connection {host['connectionName']} not in connectionReferences")

    for s in strings(definition):
        for expr in expressions(s):
            if not balanced(expr):
                err(flow, f"unbalanced expression: {expr[:120]}")
            for kind, ref in re.findall(r"\b(outputs|body|items|actions)\('([^']+)'\)", expr):
                if ref not in all_actions:
                    err(flow, f"{kind}('{ref}') refers to an unknown action")
            for ref in re.findall(r"\bvariables\('([^']+)'\)", expr):
                if ref not in variables:
                    err(flow, f"variables('{ref}') is never initialized")
            for ref in re.findall(r"\bparameters\('([^']+)'\)", expr):
                if ref not in params:
                    err(flow, f"parameters('{ref}') is not declared")

    print(f"checked {flow}: {len(all_actions)} triggers/actions, variables={sorted(variables)}")

if errors:
    print("\n".join(["", "FAILED:"] + errors))
    sys.exit(1)
print("all checks passed")
