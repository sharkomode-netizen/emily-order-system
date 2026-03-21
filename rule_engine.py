"""
Rule Engine — 加载并应用 correction_rules.json 中的声明式规则。
纯函数，零 AI 调用，零副作用。
"""
import json
import os
import re
import time
import logging

logger = logging.getLogger(__name__)

RULES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'correction_rules.json')
_rules_cache = {'mtime': 0, 'rules': []}


def load_rules():
    """Load rules from JSON, with file mtime caching."""
    global _rules_cache
    try:
        mtime = os.path.getmtime(RULES_FILE)
        if mtime == _rules_cache['mtime']:
            return _rules_cache['rules']
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        _rules_cache = {'mtime': mtime, 'rules': data.get('rules', [])}
        return _rules_cache['rules']
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def _matches_condition(rule, data, conversion_type):
    """Check if a rule's conditions match the current context."""
    # Check conversion type
    rule_conv = rule.get('conversion', 'all')
    if rule_conv != 'all' and rule_conv != conversion_type:
        return False

    # Check enabled
    if not rule.get('enabled', True):
        return False

    # Only apply verified rules in production
    if not rule.get('verified', False):
        return False

    condition = rule.get('condition', {})

    # Customer pattern matching
    if 'customer_pattern' in condition:
        customer = (data.get('customer', '') or '').lower()
        if condition['customer_pattern'].lower() not in customer:
            return False

    # Field existence check
    if 'field_exists' in condition:
        if condition['field_exists'] not in data:
            return False

    return True


# === Safe transform functions (fixed vocabulary) ===

def _transform_round(value, params):
    """Round numeric value."""
    try:
        return round(float(value), params.get('precision', 2))
    except (ValueError, TypeError):
        return value


def _transform_strip_prefix(value, params):
    """Strip leading digits/chars from a string."""
    pattern = params.get('pattern', r'^\d+')
    return re.sub(pattern, '', str(value)).strip()


def _transform_replace(value, params):
    """String replacement."""
    old = params.get('old', '')
    new = params.get('new', '')
    if old:
        return str(value).replace(old, new)
    return value


def _transform_set_default(value, params):
    """Set default value if empty."""
    if not value:
        return params.get('default', '')
    return value


def _transform_uppercase(value, params):
    return str(value).upper()


def _transform_lowercase(value, params):
    return str(value).lower()


def _transform_number_format(value, params):
    """Format number string."""
    fmt = params.get('format', '.2f')
    try:
        return format(float(value), fmt)
    except (ValueError, TypeError):
        return value


TRANSFORMS = {
    'round': _transform_round,
    'strip_prefix': _transform_strip_prefix,
    'replace': _transform_replace,
    'set_default': _transform_set_default,
    'uppercase': _transform_uppercase,
    'lowercase': _transform_lowercase,
    'number_format': _transform_number_format,
}


def apply_parse_rules(parsed_data, conversion_type):
    """Apply matching rules to parsed data dict. Returns modified data."""
    rules = load_rules()
    applied = 0

    for rule in rules:
        if rule.get('scope') != 'parse':
            continue
        if not _matches_condition(rule, parsed_data, conversion_type):
            continue

        action = rule.get('action', {})
        rule_type = rule.get('type', '')

        try:
            if rule_type == 'field_mapping':
                field = action.get('field', '')
                transform = action.get('transform', '')
                if field and transform in TRANSFORMS:
                    if field in parsed_data:
                        parsed_data[field] = TRANSFORMS[transform](parsed_data[field], action)
                    # Also apply to items
                    for item in parsed_data.get('items', []):
                        if field in item:
                            item[field] = TRANSFORMS[transform](item[field], action)
                    applied += 1

            elif rule_type == 'value_correction':
                field = action.get('field', '')
                transform = action.get('transform', '')
                if field and transform in TRANSFORMS:
                    for item in parsed_data.get('items', []):
                        if field in item:
                            item[field] = TRANSFORMS[transform](item[field], action)
                    applied += 1

            elif rule_type == 'customer_template':
                # Merge customer-specific defaults
                defaults = action.get('defaults', {})
                for k, v in defaults.items():
                    if not parsed_data.get(k):
                        parsed_data[k] = v
                applied += 1

        except Exception as e:
            logger.warning(f"Rule {rule.get('id')} failed: {e}")

    if applied > 0:
        logger.info(f"Applied {applied} correction rules for {conversion_type}")

    return parsed_data


def get_prompt_patches(conversion_type):
    """Get additional AI prompt instructions from rules."""
    rules = load_rules()
    patches = []

    for rule in rules:
        if rule.get('type') != 'prompt_patch':
            continue
        if not rule.get('enabled', True) or not rule.get('verified', False):
            continue
        rule_conv = rule.get('conversion', 'all')
        if rule_conv != 'all' and rule_conv != conversion_type:
            continue

        instruction = rule.get('action', {}).get('append_instruction', '')
        if instruction:
            patches.append(instruction)

    if patches:
        return '\n\nAdditional rules:\n' + '\n'.join(f'- {p}' for p in patches)
    return ''


def get_format_rules(conversion_type):
    """Get format override rules for generation phase."""
    rules = load_rules()
    fmt_rules = []

    for rule in rules:
        if rule.get('type') != 'format_override':
            continue
        if not rule.get('enabled', True) or not rule.get('verified', False):
            continue
        rule_conv = rule.get('conversion', 'all')
        if rule_conv != 'all' and rule_conv != conversion_type:
            continue
        fmt_rules.append(rule)

    return fmt_rules


def save_rule(rule):
    """Append a new rule to correction_rules.json."""
    try:
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {'version': 1, 'rules': []}

    # Deduplicate by checking existing rule descriptions
    for existing in data['rules']:
        if existing.get('description') == rule.get('description'):
            return False  # already exists

    data['rules'].append(rule)

    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Invalidate cache
    _rules_cache['mtime'] = 0
    return True
