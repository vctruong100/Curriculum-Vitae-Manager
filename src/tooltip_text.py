import logging

_logger = logging.getLogger(__name__)

TOOLTIP_MAX_WIDTH = 320

TOOLTIP_DEFAULT = "No description available."

TOOLTIP_TEXT = {
    "fuzzy_threshold_full": (
        "Minimum fuzzy-match score (0\u2013100) when comparing the full "
        "description of a CV study to master studies. Higher values require "
        "closer matches. Lowering this may match more studies but risks "
        "false positives."
    ),
    "fuzzy_threshold_masked": (
        "Minimum fuzzy-match score (0\u2013100) when comparing the masked "
        "(redacted) description. Used during Mode B redaction matching. "
        "Lower values are more permissive."
    ),
    "benchmark_min_count": (
        "If the number of studies in the latest year is less than or equal "
        "to this value minus one, the benchmark year steps back by one. "
        "Controls how aggressively new studies are injected."
    ),
    "highlight_inserted": (
        "When enabled, newly injected studies will have their sponsor and "
        "protocol name highlighted in yellow in the output .docx file. "
        "Existing studies are never highlighted."
    ),
    "font_name": (
        "Font family applied to all text in the output .docx. Must be one "
        "of the allowed fonts. Changing this affects all generated "
        "paragraphs including headings and study entries."
    ),
    "font_size": (
        "Font size in points (6\u201372) applied to all text in the output "
        ".docx. Affects study entries, headings, and all generated content."
    ),
    "uncategorized_label": (
        "Display label used for studies that cannot be matched to any "
        "master category. This label appears as a phase heading in the "
        "output document. Must be non-empty."
    ),
    "year_inference_full_threshold": (
        "Minimum match score for inferring a missing year from master "
        "studies using full description comparison."
    ),
    "year_inference_masked_threshold": (
        "Minimum match score for inferring a missing year from master "
        "studies using masked description comparison."
    ),
    "backup_retention_days": (
        "Number of days to retain automatic database backups. Backups "
        "older than this are deleted at startup."
    ),
    "log_retention_days": (
        "Number of days to retain operation log files. Logs older than "
        "this are deleted at startup."
    ),
    "offline_guard_enabled": (
        "When enabled, blocks all network sockets at startup and scans "
        "for proxy environment variables. Ensures the application "
        "remains fully offline."
    ),
    "enable_sort_existing": (
        "When enabled, all studies (existing and new) are sorted during "
        "Update/Inject. When disabled, only newly inserted studies are "
        "sorted; existing CV order is preserved."
    ),
    "phase_order": (
        "Comma-separated list of phases in desired display order. "
        "Controls how phases are sorted in the output document."
    ),
}


def get_tooltip_text(key):
    text = TOOLTIP_TEXT.get(key)
    if text is None:
        _logger.debug("[Tooltip] No tooltip text for key '%s', using default", key)
        return TOOLTIP_DEFAULT
    return text
