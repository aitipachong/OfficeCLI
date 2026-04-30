// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

// SCOPE: Word-only i18n write/read helpers. This file consolidates two
// duplicated patterns previously scattered across Set.cs, Set.Element.cs,
// Add.Text.cs, Add.Structure.cs, and Navigation.cs:
//
//   1) The RTL cascade — `direction=rtl` requires <w:bidi/> on pPr +
//      <w:rtl/> on the paragraph mark rPr + <w:rtl/> on every run rPr.
//      Word's UI writes all three; missing any of them produces a mixed-bidi
//      paragraph that renders incorrectly for Arabic/Hebrew fonts.
//
//   2) Complex-script (CS) run readback — font.cs / size.cs / bold.cs /
//      italic.cs were read at two sites in Navigation with subtly different
//      fallback semantics. ReadComplexScriptRunFormatting unifies them.
//
// DO NOT add: locale → font mapping (lives in Core/LocaleFontRegistry),
// HTML preview lang/CSS fallback (lives in HtmlPreview.* and there has only
// one call site each), themeFontLang stamping (lives in BlankDocCreator,
// single site). Those don't have duplication worth abstracting.
//
// Pptx/Excel handlers have similar patterns but are intentionally NOT
// covered here — wait until two handlers actually share, then promote to
// Core/. This file stays Word-only.
public partial class WordHandler
{
    /// <summary>
    /// Apply the full RTL cascade (<w:bidi/> + paragraph-mark <w:rtl/> +
    /// every run's <w:rtl/>) to <paramref name="paragraph"/>. Idempotent and
    /// reversible: pass <paramref name="rtl"/>=false to clear the cascade.
    ///
    /// <para>
    /// CONSISTENCY(rtl-cascade): a paragraph-level <w:bidi/> alone only flips
    /// layout (page side, mark anchor); it does NOT reverse the run-internal
    /// character order. Word's UI also writes <w:rtl/> on every run and on
    /// the paragraph mark when the user toggles paragraph direction — this
    /// helper mirrors that so a single direction=rtl produces a fully
    /// Arabic-correct paragraph. Used by all paragraph-level callers (Set,
    /// SetElement, Add header/footer, table cell).
    /// </para>
    ///
    /// <para>
    /// One deliberate exclusion: <c>StyleRunProperties</c> in
    /// Add.Structure.cs:498-500 stamps <w:rFonts> only and intentionally
    /// omits <w:rtl/> due to schema-order constraints there. That site stays
    /// hand-rolled — do not redirect through this helper.
    /// </para>
    /// </summary>
    private void ApplyDirectionCascade(Paragraph paragraph, bool rtl)
    {
        var pProps = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());

        if (rtl)
            pProps.BiDi = new BiDi();
        else
            pProps.RemoveAllChildren<BiDi>();

        var markRPr = pProps.ParagraphMarkRunProperties
            ?? pProps.AppendChild(new ParagraphMarkRunProperties());
        ApplyRunFormatting(markRPr, "direction", rtl ? "rtl" : "ltr");

        foreach (var run in paragraph.Descendants<Run>())
        {
            var rPr = EnsureRunProperties(run);
            ApplyRunFormatting(rPr, "direction", rtl ? "rtl" : "ltr");
        }
    }

    /// <summary>
    /// Read complex-script run formatting (<w:rFonts cs/>, <w:szCs/>,
    /// <w:bCs/>, <w:iCs/>) into <paramref name="format"/>. Mirrors the
    /// canonical keys font.cs / size.cs / bold.cs / italic.cs.
    ///
    /// <para>
    /// Two-arg form lets the paragraph readback site fall back from the
    /// first run's rPr to the paragraph-mark rPr (covers paragraphs that
    /// have CS flags on the mark but no runs yet). Run-level callers pass
    /// <paramref name="fallback"/>=null.
    /// </para>
    ///
    /// <para>
    /// Skips keys that already exist in <paramref name="format"/> so callers
    /// can layer this on top of other readers without overwriting.
    /// </para>
    /// </summary>
    private static void ReadComplexScriptRunFormatting(
        OpenXmlCompositeElement? primary,
        OpenXmlCompositeElement? fallback,
        IDictionary<string, object?> format)
    {
        // font.cs — only set by ApplyRunFormatting; falls under <w:rFonts>.
        var rFontsP = primary?.GetFirstChild<RunFonts>();
        var rFontsF = fallback?.GetFirstChild<RunFonts>();
        var fontCs = !string.IsNullOrEmpty(rFontsP?.ComplexScript?.Value)
            ? rFontsP!.ComplexScript!.Value
            : (!string.IsNullOrEmpty(rFontsF?.ComplexScript?.Value)
                ? rFontsF!.ComplexScript!.Value
                : null);
        if (fontCs != null && !format.ContainsKey("font.cs"))
            format["font.cs"] = fontCs;

        // size.cs — half-points, formatted as "Npt".
        var szCsEl = primary?.GetFirstChild<FontSizeComplexScript>()
            ?? fallback?.GetFirstChild<FontSizeComplexScript>();
        if (szCsEl?.Val?.Value is string szCsVal
            && int.TryParse(szCsVal, out var szCsHalfPt)
            && !format.ContainsKey("size.cs"))
        {
            format["size.cs"] = $"{szCsHalfPt / 2.0:0.##}pt";
        }

        // bold.cs / italic.cs — boolean flags.
        var bCsEl = primary?.GetFirstChild<BoldComplexScript>()
            ?? fallback?.GetFirstChild<BoldComplexScript>();
        if (bCsEl != null && !format.ContainsKey("bold.cs"))
            format["bold.cs"] = true;

        var iCsEl = primary?.GetFirstChild<ItalicComplexScript>()
            ?? fallback?.GetFirstChild<ItalicComplexScript>();
        if (iCsEl != null && !format.ContainsKey("italic.cs"))
            format["italic.cs"] = true;
    }
}
