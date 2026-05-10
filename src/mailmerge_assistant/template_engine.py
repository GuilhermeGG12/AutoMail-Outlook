from __future__ import annotations

import re

PLACEHOLDER_RE = re.compile(r"{([^{}]+)}")


class TemplateRenderError(ValueError):
    """Raised when a template cannot be rendered safely."""


def find_placeholders(template: str) -> set[str]:
    return set(PLACEHOLDER_RE.findall(template))


def render_template(template: str, values: dict[str, str]) -> str:
    missing = sorted(name for name in find_placeholders(template) if name not in values)
    if missing:
        missing_text = ", ".join(missing)
        raise TemplateRenderError(f"Variáveis ausentes no modelo: {missing_text}")

    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        return values[key]

    return PLACEHOLDER_RE.sub(replace, template)
