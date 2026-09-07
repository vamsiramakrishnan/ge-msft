# GENERATED from skill/shared/language_manifest.py by emit:language.
"""Canonical standalone manifest reader and structural guard evaluator, copied by emit:language.

Missing, malformed or incompatible generated data is a packaging error. Never infer a permissive
language from an incomplete bundle. No live host values or approval are evaluated here.
"""

import json
import math
from pathlib import Path


class ManifestError(ValueError):
    pass


def _strings(value, label):
    if not isinstance(value, list) or not value or any(not isinstance(v, str) or not v for v in value) or len(value) != len(set(value)):
        raise ManifestError(f"Invalid {label}: expected a nonempty string list")


def _reject_constant(value):
    raise ManifestError(f"Non-JSON numeric constant {value}")


def load_manifest(path, version):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"), parse_constant=_reject_constant)
        if not isinstance(data, dict) or data.get("version") != version:
            raise ManifestError(f"Expected {version}")
        if version == "m365-plan/1.0":
            for key in ("scalarKeys", "listKeys", "brackets", "intents", "surfaces", "confidence", "contextHints"):
                _strings(data.get(key), key)
        elif version == "m365-cli/1.0":
            for key in ("read", "workspace", "control", "write"):
                _strings(data["verbs"].get(key), f"verbs.{key}")
            for key in ("transforms", "effectVerbs", "actuationKinds"):
                _strings(data.get(key), key)
            if not isinstance(data.get("commandHelp"), dict) or not isinstance(data.get("capabilityRegistry"), list):
                raise ManifestError("Missing generated help or capability registry")
            meta = data["preflight"]
            if meta.get("formatVersion") != 1:
                raise ManifestError("Unsupported preflight metadata version")
            for key in ("contextKinds", "contextHints", "analysisBindingKinds", "readPhaseVerbs"):
                _strings(meta.get(key), key)
            for key in ("approvalByKind", "approvalByVerb"):
                if not isinstance(meta.get(key), dict) or not meta[key]:
                    raise ManifestError(f"Missing {key}")
                for authority in meta[key].values():
                    if (authority.get("approvalClass") not in ("in-document", "external", "estate", "irreversible")
                            or type(authority.get("reversible")) is not bool):
                        raise ManifestError(f"Invalid {key} authority")
            if set(meta["approvalByKind"]) != set(data["actuationKinds"]):
                raise ManifestError("Approval metadata does not cover the actuation catalogue")
            if not (set(data["verbs"]["write"]) | {"share"}) <= set(meta["approvalByVerb"]):
                raise ManifestError("Approval metadata does not cover write verbs")
            if not isinstance(meta.get("analysisGuards"), dict) or not meta["analysisGuards"]:
                raise ManifestError("Missing analysis guard metadata")
            for fields in meta["analysisGuards"].values():
                if not isinstance(fields, dict) or not fields:
                    raise ManifestError("Missing analysis field guards")
                for descriptor in fields.values():
                    validate_descriptor(descriptor)
        else:
            raise ManifestError("Unsupported manifest version")
        return data
    except (OSError, ValueError, KeyError, TypeError, AttributeError) as error:
        raise ManifestError(f"Cannot load bundled {Path(path).name}: {error}. Rebuild with skills:generate and package the complete skill.") from error


def validate_descriptor(schema):
    if not isinstance(schema, dict) or schema.get("type") not in ("object", "array", "string", "number", "boolean"):
        raise ManifestError("Unsupported generated guard descriptor")
    if "optional" in schema and type(schema["optional"]) is not bool:
        raise ManifestError("Invalid optional guard metadata")
    allowed = {"type", "optional"} | {
        "object": {"properties", "strict"}, "array": {"items", "min", "max"},
        "string": {"checks"}, "number": {"checks"}, "boolean": set(),
    }[schema["type"]]
    if set(schema) - allowed:
        raise ManifestError("Unsupported generated guard fields")
    if schema["type"] == "object":
        if not isinstance(schema.get("properties"), dict) or type(schema.get("strict")) is not bool:
            raise ManifestError("Invalid object guard metadata")
        for child in schema["properties"].values():
            validate_descriptor(child)
    if schema["type"] == "array":
        validate_descriptor(schema["items"])
    for boundary in ("min", "max"):
        if boundary in schema and (type(schema[boundary]) is not int or schema[boundary] < 0):
            raise ManifestError("Invalid array guard bounds")
    if not isinstance(schema.get("checks", []), list):
        raise ManifestError("Invalid guard checks")
    for check in schema.get("checks", []):
        if not isinstance(check, dict):
            raise ManifestError("Invalid guard check")
        allowed_checks = ("min", "max", "int", "finite") if schema["type"] == "number" else ("min", "max")
        if check.get("kind") not in allowed_checks or set(check) - {"kind", "value", "inclusive"}:
            raise ManifestError("Unsupported generated guard check")
        bound = check.get("value")
        if check["kind"] in ("min", "max") and (type(bound) not in (int, float)
                or type(bound) is float and not math.isfinite(bound)):
            raise ManifestError("Invalid guard bound")
        if "inclusive" in check and type(check["inclusive"]) is not bool:
            raise ManifestError("Invalid guard boundary metadata")


def guard_errors(value, schema, path):
    kind = schema["type"]
    types = {"object": dict, "array": list, "string": str, "boolean": bool}
    valid = (type(value) is int or type(value) is float and math.isfinite(value)) if kind == "number" else isinstance(value, types[kind])
    if not valid:
        return [f"{path} must be {kind}"]
    errors = []
    if kind == "object":
        properties = schema["properties"]
        if schema.get("strict") and set(value) - set(properties):
            errors.append(f"{path} has unknown fields")
        for key, child in properties.items():
            if key in value:
                errors.extend(guard_errors(value[key], child, f"{path}.{key}"))
            elif not child.get("optional"):
                errors.append(f"{path}.{key} is required")
    elif kind == "array":
        if "min" in schema and len(value) < schema["min"] or "max" in schema and len(value) > schema["max"]:
            errors.append(f"{path} exceeds its item bounds")
        for index, item in enumerate(value):
            errors.extend(guard_errors(item, schema["items"], f"{path}[{index}]"))
    for check in schema.get("checks", []):
        # Zod uses JavaScript UTF-16 length, including two code units for supplementary characters.
        observed = len(value.encode("utf-16-le", errors="surrogatepass")) // 2 if kind == "string" else value
        if check["kind"] == "int" and observed != int(observed):
            errors.append(f"{path} must be an integer")
        if (check["kind"] == "min" and (observed < check["value"] or observed == check["value"] and check.get("inclusive") is False)
                or check["kind"] == "max" and (observed > check["value"] or observed == check["value"] and check.get("inclusive") is False)):
            errors.append(f"{path} exceeds its bounds")
    return errors
