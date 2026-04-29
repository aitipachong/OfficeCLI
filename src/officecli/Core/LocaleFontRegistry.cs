// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Locale → default font mapping for fresh blank documents. Mirrors the
/// data-driven approach LibreOffice uses (VCL.xcu): given a locale tag, pick
/// reasonable defaults for the Latin / EastAsian / ComplexScript font slots.
///
/// We deliberately keep this small (one line per locale family) rather than
/// trying to model every Office localization. When no locale is supplied,
/// returning all-empty values lets the host application substitute its own
/// UI-locale defaults — that's the POI-aligned behaviour BlankDocCreator
/// already had after we removed the "宋体" hardcode.
///
/// Font names are chosen for cross-platform availability (typefaces shipped
/// by Microsoft Office on Windows and macOS, plus Apple Sans equivalents).
/// </summary>
public static class LocaleFontRegistry
{
    /// <summary>
    /// Resolve a locale tag (e.g. "zh-CN", "ja", "ar-SA") to a per-script
    /// font triple. Returns (null, null, null) when no locale is supplied
    /// or the tag is unknown — callers should treat that as "leave the
    /// docDefaults blank, let the host application decide".
    /// </summary>
    public static (string? Latin, string? EastAsia, string? ComplexScript) Resolve(string? locale)
    {
        if (string.IsNullOrWhiteSpace(locale)) return (null, null, null);

        // Match on language-only first; full tag lookups (e.g. zh-Hant) are
        // routed through the language-only entry unless a region-specific
        // variant exists.
        var lower = locale.Replace('_', '-').ToLowerInvariant();
        var lang = lower.Split('-')[0];

        // Fully-tagged regional variants take precedence.
        switch (lower)
        {
            case "zh-tw" or "zh-hk" or "zh-mo" or "zh-hant":
                return ("Times New Roman", "新細明體", null);
            case "zh-cn" or "zh-sg" or "zh-hans":
                return ("Times New Roman", "等线", null);
        }

        // Language-only fall-throughs.
        return lang switch
        {
            "zh" => ("Times New Roman", "等线", null),
            "ja" => ("Times New Roman", "游明朝", null),
            "ko" => ("Times New Roman", "맑은 고딕", null),
            "ar" => ("Times New Roman", null, "Arabic Typesetting"),
            "he" => ("Times New Roman", null, "Times New Roman"),
            "th" => ("Times New Roman", null, "Tahoma"),
            "fa" => ("Times New Roman", null, "B Nazanin"),
            "ur" => ("Times New Roman", null, "Jameel Noori Nastaleeq"),
            "hi" => ("Times New Roman", null, "Mangal"),
            "en" or "fr" or "de" or "es" or "it" or "pt" or "nl" or "ru" or "pl"
                => ("Times New Roman", null, null),
            _ => (null, null, null)
        };
    }
}
