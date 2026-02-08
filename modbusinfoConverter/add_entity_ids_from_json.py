#!/usr/bin/env python3
"""Add entity_id field to JSON configs using English names from existing files.

This script is useful when Excel source files are not available but both
German and English JSON configs exist.
"""

import json
import re
from pathlib import Path


# Character replacements for entity ID sanitization (same as coordinator.py)
_ENTITY_ID_REPLACEMENTS = str.maketrans({
    " ": "_",
    ".": "_",
    "/": "_",
    "-": "_",
    ":": "_",
    "ä": "ae",
    "ö": "oe",
    "ü": "ue",
    "Ä": "ae",
    "Ö": "oe",
    "Ü": "ue",
    "ß": "ss",
    "&": "and",
    "@": "at",
    "(": "",
    ")": "",
    "#": "",
    "!": "",
    "?": "",
    ",": "",
    ";": "",
    "'": "",
    '"': "",
})


def sanitize_for_entity_id(name: str) -> str:
    """Sanitize a string for use in Home Assistant entity IDs."""
    if not name:
        return ""

    result = name.translate(_ENTITY_ID_REPLACEMENTS)
    result = result.lower()
    result = re.sub(r'[^a-z0-9_]', '', result)
    result = re.sub(r'_+', '_', result)
    result = result.strip('_')

    return result


def load_english_names_from_json(en_dir: Path) -> dict[int, str]:
    """Load English register names from JSON files.

    Returns a mapping of starting_address -> English name.
    """
    english_names: dict[int, str] = {}

    # Process all JSON files in the English directory
    for json_file in en_dir.rglob("*.json"):
        if json_file.name in ["value_tables.json", "alarm_codes.json"]:
            continue  # Skip non-register files

        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Handle different JSON structures
            registers = []
            if "universal_registers" in data:
                registers = data["universal_registers"]
            elif "registers" in data:
                registers = data["registers"]

            for reg in registers:
                address = reg.get("starting_address")
                name = reg.get("name")
                if address is not None and name:
                    english_names[int(address)] = str(name).strip()

        except Exception as e:
            print(f"  Warning: Could not process {json_file}: {e}")

    return english_names


def add_entity_ids_to_registers(registers: list[dict], english_names: dict[int, str]) -> int:
    """Add entity_id field to registers list.

    Returns count of registers updated.
    """
    updated = 0
    for reg in registers:
        if "entity_id" in reg:
            continue  # Already has entity_id

        address = reg.get("starting_address")
        if address is None:
            continue

        # Get English name, fall back to current name
        english_name = english_names.get(int(address), reg.get("name", ""))
        entity_id = sanitize_for_entity_id(english_name)

        if entity_id:
            # Insert entity_id after name for consistent ordering
            new_reg = {}
            for key, value in reg.items():
                new_reg[key] = value
                if key == "name":
                    new_reg["entity_id"] = entity_id
            reg.clear()
            reg.update(new_reg)
            updated += 1

    return updated


def process_json_file(json_file: Path, english_names: dict[int, str]) -> bool:
    """Process a single JSON file to add entity_id fields.

    Returns True if file was modified.
    """
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        modified = False

        # Handle different JSON structures
        if "universal_registers" in data:
            count = add_entity_ids_to_registers(data["universal_registers"], english_names)
            if count > 0:
                modified = True
                print(f"    Updated {count} universal registers")

        if "registers" in data:
            count = add_entity_ids_to_registers(data["registers"], english_names)
            if count > 0:
                modified = True
                print(f"    Updated {count} registers")

        if modified:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True

    except Exception as e:
        print(f"  Error processing {json_file}: {e}")

    return False


def process_version(version_dir: Path) -> None:
    """Process a version directory to add entity_id fields."""
    en_dir = version_dir / "en"
    de_dir = version_dir / "de"

    if not en_dir.exists():
        print(f"  No English directory found, skipping")
        return

    print(f"  Loading English names from {en_dir}...")
    english_names = load_english_names_from_json(en_dir)
    print(f"  Loaded {len(english_names)} English names")

    # Process both language directories
    for lang_dir in [de_dir, en_dir]:
        if not lang_dir.exists():
            continue

        print(f"  Processing {lang_dir.name}/...")

        for json_file in lang_dir.rglob("*.json"):
            if json_file.name in ["value_tables.json", "alarm_codes.json"]:
                continue

            relative = json_file.relative_to(lang_dir)
            print(f"    {relative}...")
            process_json_file(json_file, english_names)


def main():
    """Main entry point."""
    import sys

    # Default to custom_components config directory
    script_dir = Path(__file__).parent.parent
    config_dir = script_dir / "custom_components" / "kwb_heating" / "config" / "versions"

    # Allow specifying a specific version
    if len(sys.argv) > 1:
        version = sys.argv[1]
        version_dirs = [config_dir / f"v{version}"]
    else:
        # Process all versions
        version_dirs = sorted(config_dir.glob("v*"))

    print("Adding entity_id fields to JSON configs")
    print(f"Config directory: {config_dir}")
    print("=" * 60)

    for version_dir in version_dirs:
        if not version_dir.is_dir():
            continue

        print(f"\nProcessing {version_dir.name}...")
        process_version(version_dir)

    print("\n" + "=" * 60)
    print("Done!")


if __name__ == "__main__":
    main()
