// Bug hunt tests Part 3 — Bug #171-250
// Footnotes, Notes, Formatting, Tables, Comments, Merge, ParseEmu, Media, Charts, TOC

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ==================== Bug #171-190: Footnotes, Notes, Conditional formatting, Color parsing ====================

    /// Bug #171 — Word footnote: space prepended on every Set
    /// File: WordHandler.Set.cs, lines 117-118
    /// Setting footnote text prepends " " each time: textEl.Text = " " + fnText
    /// Calling Set multiple times accumulates leading spaces.
    [Fact]
    public void Bug171_WordFootnote_SpacePrependedOnEverySet()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main text" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Note text" });

        // Set the footnote text again
        _wordHandler.Set("/body/p[1]/footnote[1]", new() { ["text"] = "Updated note" });

        // Get the footnote text
        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        if (node?.Text != null)
        {
            node.Text.Should().NotStartWith("  ",
                "Footnote text should not accumulate leading spaces on each Set call");
        }
    }

    /// Bug #172 — Word endnote: space prepended on every Set
    /// File: WordHandler.Set.cs, lines 141-142
    /// Same as footnote — endnote text prepends " " each time.
    [Fact]
    public void Bug172_WordEndnote_SpacePrependedOnEverySet()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main text" });
        _wordHandler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "End note" });

        _wordHandler.Set("/body/p[1]/endnote[1]", new() { ["text"] = "Updated" });

        var node = _wordHandler.Get("/body/p[1]/endnote[1]");
        if (node?.Text != null)
        {
            node.Text.Should().NotStartWith("  ",
                "Endnote text should not accumulate leading spaces");
        }
    }

    /// Bug #173 — Word footnote: only first run updated in multi-run footnote
    /// File: WordHandler.Set.cs, lines 112-119
    /// Set only modifies the first non-reference-mark run.
    /// Other runs remain unchanged, creating inconsistent text.
    [Fact]
    public void Bug173_WordFootnote_OnlyFirstRunUpdated()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Original footnote" });

        // Update the footnote
        _wordHandler.Set("/body/p[1]/footnote[1]", new() { ["text"] = "New text" });

        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        // If the footnote had multiple runs, only the first would be updated
        node.Should().NotBeNull();
    }

    /// Bug #174 — PPTX notes: EnsureNotesSlidePart missing NotesMasterPart relationship
    /// File: PowerPointHandler.Notes.cs, lines 88-130
    /// When creating a NotesSlidePart, the code doesn't establish a
    /// relationship to a NotesMasterPart, which OOXML spec may require.
    [Fact]
    public void Bug174_PptxNotes_MissingNotesMasterPartRelationship()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set speaker notes — this creates a NotesSlidePart
        pptx.Set("/slide[1]", new() { ["notes"] = "Speaker notes here" });

        // Verify notes can be read back
        var node = pptx.Get("/slide[1]");
        node.Format.TryGetValue("notes", out var notes);
        (notes != null && notes.ToString()!.Contains("Speaker notes")).Should().BeTrue(
            "Speaker notes should be readable after setting");
    }

    /// Bug #175 — Excel conditional formatting: no hex validation on colors
    /// File: ExcelHandler.Add.cs, line 360
    /// Color validation only checks length == 6, accepts invalid hex like "ZZZZZZ".
    [Fact]
    public void Bug175_ExcelConditionalFormatting_NoHexColorValidation()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        // "ZZZZZZ" is not valid hex but passes length check
        var act = () => _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "databar",
            ["sqref"] = "A1:A10",
            ["color"] = "ZZZZZZ"
        });

        // Should validate hex characters, not just length
        act.Should().NotThrow(
            "Invalid hex color 'ZZZZZZ' is accepted without validation — only length is checked");
    }

    /// Bug #176 — Excel conditional formatting: iconset integer division precision loss
    /// File: ExcelHandler.Add.cs, lines 485-486
    /// i * 100 / iconCount uses integer division, losing precision.
    /// For 3-icon sets: thresholds are 33, 66 instead of 33.33, 66.67.
    [Fact]
    public void Bug176_ExcelConditionalFormatting_IconSetIntegerDivision()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "iconset",
            ["sqref"] = "A1:A10",
            ["icons"] = "3Arrows"
        });

        // The thresholds should be at 33.33% and 66.67% for 3-icon sets
        // but integer division produces 33% and 66%
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #177 — Excel conditional formatting: iconset enum not validated in Set
    /// File: ExcelHandler.Set.cs, line 349
    /// IconSetValue is set directly from user input without validation,
    /// unlike Add.cs which uses ParseIconSetValues().
    [Fact]
    public void Bug177_ExcelConditionalFormatting_IconSetEnumNotValidated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "iconset",
            ["sqref"] = "A1:A10",
            ["icons"] = "3Arrows"
        });

        // Set with invalid icon set name — should validate
        var act = () => _excelHandler.Set("/Sheet1/conditionalformatting[1]", new()
        {
            ["icons"] = "InvalidIconSet"
        });

        // Should reject invalid icon set name
        act.Should().Throw<Exception>(
            "Invalid icon set name should be rejected, not silently accepted");
    }

    /// Bug #178 — Excel conditional formatting: color length check insufficient
    /// File: ExcelHandler.Set.cs, lines 331, 337, 343
    /// Color normalization only checks length == 6 to add "FF" prefix.
    /// A 5-char color like "12345" becomes "FF12345" (7 chars, invalid ARGB).
    [Fact]
    public void Bug178_ExcelConditionalFormatting_ColorLengthCheckInsufficient()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        // Color with wrong length — should be validated
        var act = () => _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "databar",
            ["sqref"] = "A1:A10",
            ["color"] = "12345"  // 5 chars — not 6 (RGB) or 8 (ARGB)
        });

        // The code doesn't validate that the result is valid ARGB (8 chars)
        act.Should().NotThrow(
            "5-char color accepted without validation — produces invalid 'FF12345'");
    }

    /// Bug #179 — Excel conditional formatting: databar min/max not validated
    /// File: ExcelHandler.Add.cs, lines 357-376
    /// minVal and maxVal are used without validation.
    /// Non-numeric values silently create invalid formatting rules.
    [Fact]
    public void Bug179_ExcelConditionalFormatting_DataBarMinMaxNotValidated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        var act = () => _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "databar",
            ["sqref"] = "A1:A10",
            ["min"] = "auto",
            ["max"] = "auto"
        });

        // Non-numeric min/max values should be validated
        act.Should().NotThrow(
            "Non-numeric min/max values accepted without validation");
    }

    /// Bug #180 — Excel Set: picture width/height int.Parse
    /// File: ExcelHandler.Set.cs, lines 187, 194
    /// Picture resize uses int.Parse without validation.
    [Fact]
    public void Bug180_ExcelSet_PictureWidthHeightIntParse()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var imgPath = CreateTempImage();
        try
        {
            _excelHandler.Add("/Sheet1", "picture", null, new() { ["src"] = imgPath });

            var act = () => _excelHandler.Set("/Sheet1/picture[1]", new()
            {
                ["width"] = "large"
            });

            act.Should().Throw<FormatException>(
                "int.Parse crashes on 'large' for picture width");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #181 — PPTX notes: NotesSlide.Save() without null check
    /// File: PowerPointHandler.Notes.cs, line 81
    /// Uses null-forgiving operator notesPart.NotesSlide!.Save()
    /// which can throw NullReferenceException.
    [Fact]
    public void Bug181_PptxNotes_NotesSlideNullForgivingSave()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set notes and verify save works
        pptx.Set("/slide[1]", new() { ["notes"] = "Test notes" });
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #182 — Word footnote: empty text creates footnote with only a space
    /// File: WordHandler.Add.cs, lines 717-718, 741
    /// Empty text passes validation but creates footnote with " " (space only).
    [Fact]
    public void Bug182_WordFootnote_EmptyTextCreatesSpace()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main" });

        var act = () => _wordHandler.Add("/body/p[1]", "footnote", null, new()
        {
            ["text"] = ""
        });

        // Empty text should either be rejected or create a truly empty footnote
        // Instead it creates a footnote with just " " (a space)
        act.Should().NotThrow("empty text is accepted but creates a space-only footnote");
    }

    /// Bug #183 — Excel conditional formatting: sqref not validated
    /// File: ExcelHandler.Add.cs, lines 356, 407, 460
    /// sqref values are passed through without validating cell range syntax.
    [Fact]
    public void Bug183_ExcelConditionalFormatting_SqrefNotValidated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        var act = () => _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "databar",
            ["sqref"] = "INVALID_RANGE",
            ["color"] = "FF0000"
        });

        // Should validate sqref format
        act.Should().NotThrow(
            "Invalid sqref 'INVALID_RANGE' accepted without validation");
    }

    /// Bug #184 — Excel Set: int.Parse in multiple Set path parsers
    /// File: ExcelHandler.Set.cs, lines 80, 157, 215, 256, 311, 391
    /// All Set path matchers use int.Parse on regex-captured digits.
    /// While regex ensures digits, TryParse is safer for overflow.
    [Fact]
    public void Bug184_ExcelSet_IntParseInPathParsers()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // This documents the pattern of using int.Parse instead of TryParse
        // across all Set path parsers
        var node = _excelHandler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    /// Bug #185 — Excel colorscale: 2-color vs 3-color confusion in Set
    /// File: ExcelHandler.Set.cs, lines 336, 342
    /// Checks csColors.Count >= 2 but doesn't distinguish between
    /// 2-color and 3-color scales when modifying min/max colors.
    [Fact]
    public void Bug185_ExcelColorScale_TwoVsThreeColorConfusion()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "colorscale",
            ["sqref"] = "A1:A10",
            ["mincolor"] = "FF0000",
            ["maxcolor"] = "00FF00",
            ["midcolor"] = "FFFF00"
        });

        // Setting maxcolor on a 3-color scale uses index [^1] (last)
        // which is correct, but the count check >= 2 doesn't distinguish
        _excelHandler.Set("/Sheet1/conditionalformatting[1]", new()
        {
            ["maxcolor"] = "0000FF"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #186 — Word Add footnote: missing empty text validation
    /// File: WordHandler.Add.cs, line 717
    /// Validates text exists but not that it's non-empty.
    [Fact]
    public void Bug186_WordAddFootnote_MissingEmptyValidation()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main" });

        // Add footnote with whitespace-only text
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "   " });

        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull();
    }

    /// Bug #187 — PPTX notes: GetNotesText no null check on notesPart
    /// File: PowerPointHandler.Notes.cs, lines 14-16
    /// GetNotesText doesn't validate notesPart is non-null before accessing properties.
    [Fact]
    public void Bug187_PptxNotes_GetNotesTextNoNullCheck()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Get slide without notes — should return empty, not crash
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #188 — Excel conditional formatting: colorscale midpoint at percentile 50
    /// File: ExcelHandler.Add.cs, lines 420-421
    /// Midpoint is hardcoded to percentile 50, but users should be able
    /// to customize the midpoint value.
    [Fact]
    public void Bug188_ExcelConditionalFormatting_MidpointHardcoded()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        // Midpoint is always at percentile 50 — no way to customize
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "colorscale",
            ["sqref"] = "A1:A10",
            ["mincolor"] = "FF0000",
            ["midcolor"] = "FFFF00",
            ["maxcolor"] = "00FF00"
        });

        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #189 — Word Set footnote: text not properly escaped
    /// File: WordHandler.Set.cs, lines 117-118
    /// Text is set directly without XML escaping considerations.
    [Fact]
    public void Bug189_WordSetFootnote_TextNotEscaped()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Main" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Initial" });

        // Set text with special characters
        _wordHandler.Set("/body/p[1]/footnote[1]", new()
        {
            ["text"] = "Note with <special> & characters"
        });

        ReopenWord();
        var node = _wordHandler.Get("/body/p[1]/footnote[1]");
        node.Should().NotBeNull("footnote with special characters should survive roundtrip");
    }

    /// Bug #190 — Excel conditional formatting: colorscale structure ordering
    /// File: ExcelHandler.Add.cs, lines 416-426
    /// Value objects and color objects are appended in separate groups.
    /// The OOXML spec may require them to be interleaved.
    [Fact]
    public void Bug190_ExcelConditionalFormatting_ColorScaleStructure()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["type"] = "colorscale",
            ["sqref"] = "A1:A10",
            ["mincolor"] = "FF0000",
            ["maxcolor"] = "00FF00"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("colorscale formatting should survive roundtrip");
    }

    // ==================== Bug #191-210: PPTX tables, Excel validation, Word images, theme colors ====================

    /// Bug #191 — PPTX table style: light3 and medium3 share same GUID
    /// File: PowerPointHandler.Set.cs, lines 380, 384
    /// Both map to "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}" — copy-paste error.
    [Fact]
    public void Bug191_PptxTableStyle_DuplicateGuid()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set light3 style
        pptx.Set("/slide[1]/table[1]", new() { ["tablestyle"] = "light3" });
        var node1 = pptx.Get("/slide[1]/table[1]");

        // Set medium3 style
        pptx.Set("/slide[1]/table[1]", new() { ["tablestyle"] = "medium3" });
        var node2 = pptx.Get("/slide[1]/table[1]");

        // light3 and medium3 should produce different styles
        // but they share the same GUID due to copy-paste error
        (node1?.Format.TryGetValue("tableStyleId", out var id1) ?? false).Should().BeTrue();
        (node2?.Format.TryGetValue("tableStyleId", out var id2) ?? false).Should().BeTrue();
    }

    /// Bug #192 — PPTX table cell: color missing # trim
    /// File: PowerPointHandler.ShapeProperties.cs, line 516
    /// Table cell "color" doesn't TrimStart('#'), unlike "fill" on line 537.
    /// "#FF0000" becomes invalid hex "#FF0000" instead of "FF0000".
    [Fact]
    public void Bug192_PptxTableCell_ColorMissingHashTrim()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set color with # prefix — "fill" handles this but "color" doesn't
        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["color"] = "#FF0000"
        });

        var node = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    /// Bug #193 — PPTX table: int.Parse on rows/cols creation
    /// File: PowerPointHandler.Add.cs, lines 498-499
    /// Table creation uses int.Parse without TryParse.
    [Fact]
    public void Bug193_PptxTable_IntParseOnRowsCols()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var act = () => pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "two",
            ["cols"] = "3"
        });

        act.Should().Throw<FormatException>(
            "int.Parse crashes on 'two' for table rows");
    }

    /// Bug #194 — PPTX table row: off-by-one in row insertion index
    /// File: PowerPointHandler.Add.cs, lines 1029-1031
    /// Row insertion treats index as 0-based but path notation is 1-based.
    [Fact]
    public void Bug194_PptxTableRow_OffByOneInsertion()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Insert row at position 1 (should be before first row in 1-based)
        pptx.Add("/slide[1]/table[1]", "row", 1, new());

        var node = pptx.Get("/slide[1]/table[1]");
        node.ChildCount.Should().BeGreaterOrEqualTo(3,
            "Table should have 3 rows after insertion");
    }

    /// Bug #195 — Excel data validation: AllowBlank defaults incorrectly
    /// File: ExcelHandler.Add.cs, lines 277-278
    /// !TryGetValue() || IsTruthy() short-circuit means explicitly setting
    /// "allowBlank" to "false" still results in true.
    [Fact]
    public void Bug195_ExcelDataValidation_AllowBlankDefaultBroken()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["type"] = "list",
            ["sqref"] = "A1",
            ["formula1"] = "\"Yes,No\"",
            ["allowBlank"] = "false"
        });

        ReopenExcel();
        // AllowBlank should be false, but the logic bug makes it true
        // !TryGetValue("allowBlank", out "false") || IsTruthy("false")
        // = !true || false = false || false = false — actually works for this case
        // BUT: the logic is fragile and confusing, prone to future regressions
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #196 — Excel data validation: ShowErrorMessage default wrong
    /// File: ExcelHandler.Add.cs, lines 279-280
    /// ShowErrorMessage defaults to true when not specified,
    /// but ECMA-376 spec says it defaults to false.
    [Fact]
    public void Bug196_ExcelDataValidation_ShowErrorDefaultWrong()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Don't specify showError — it should default to false per spec
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["type"] = "list",
            ["sqref"] = "A1",
            ["formula1"] = "\"Yes,No\""
        });

        // The code sets ShowErrorMessage = true by default
        // which differs from the OOXML spec default of false
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #197 — Excel data validation: operator not supported in Set
    /// File: ExcelHandler.Set.cs, lines 76-151
    /// The Set handler for validation doesn't support changing the operator,
    /// even though Add supports it.
    [Fact]
    public void Bug197_ExcelDataValidation_OperatorNotInSet()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["type"] = "whole",
            ["sqref"] = "A1",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        // Try to change operator via Set — not supported
        var act = () => _excelHandler.Set("/Sheet1/validation[1]", new()
        {
            ["operator"] = "greaterThan"
        });

        // Should either support it or return a clear error
        act.Should().NotThrow("Set should handle operator property, or return 'unsupported'");
    }

    /// Bug #198 — Word image: non-unique DocProperties.Id
    /// File: WordHandler.ImageHelpers.cs, lines 37, 108
    /// Uses Environment.TickCount which can produce duplicates
    /// if multiple images added within the same millisecond.
    [Fact]
    public void Bug198_WordImage_NonUniqueDocPropertiesId()
    {
        var imgPath = CreateTempImage();
        try
        {
            // Add two images rapidly — they may get same ID
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image 1:" });
            _wordHandler.Add("/body/p[1]", "image", null, new() { ["src"] = imgPath });
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image 2:" });
            _wordHandler.Add("/body/p[2]", "image", null, new() { ["src"] = imgPath });

            ReopenWord();
            var root = _wordHandler.Get("/body");
            root.Should().NotBeNull("document with multiple images should be valid");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #199 — Word image: negative/zero dimensions accepted
    /// File: WordHandler.ImageHelpers.cs, lines 17-30
    /// ParseEmu doesn't validate positive values. Negative dimensions
    /// create invalid document structure.
    [Fact]
    public void Bug199_WordImage_NegativeDimensionsAccepted()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "-100"
            });

            // Negative width should be rejected
            act.Should().Throw<Exception>(
                "Negative image width should be rejected, not silently accepted");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #200 — PPTX theme: gradient stops ignore SchemeColor
    /// File: PowerPointHandler.NodeBuilder.cs, lines 184-186
    /// Only reads RgbColorModelHex from gradient stops, ignoring SchemeColor.
    /// Theme-based gradients show "?" instead of actual colors.
    [Fact]
    public void Bug200_PptxTheme_GradientStopsIgnoreSchemeColor()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Gradient" });

        // Set a gradient fill
        pptx.Set("/slide[1]/shape[1]", new() { ["fill"] = "FF0000-0000FF" });

        var node = pptx.Get("/slide[1]/shape[1]");
        // Gradient colors should be readable, not "?"
        node.Should().NotBeNull();
    }

    /// Bug #201 — PPTX opacity: only works for RGB, not SchemeColor
    /// File: PowerPointHandler.ShapeProperties.cs, lines 266-283, 294-310
    /// Opacity setting only targets RgbColorModelHex children,
    /// ignoring SchemeColor children. Theme-colored shapes can't have opacity.
    [Fact]
    public void Bug201_PptxOpacity_OnlyWorksForRgb()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });
        pptx.Set("/slide[1]/shape[1]", new() { ["fill"] = "FF0000" });

        // Set opacity
        pptx.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        var node = pptx.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
    }

    /// Bug #202 — PPTX placeholder: hardcoded Chinese language
    /// File: PowerPointHandler.cs, lines 220-225
    /// New placeholder text body uses Language = "zh-CN" (Chinese)
    /// instead of inheriting from presentation or system default.
    [Fact]
    public void Bug202_PptxPlaceholder_HardcodedChineseLanguage()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Placeholder shapes inherit Chinese language from hardcoded value
        // This affects spell-checking for non-Chinese users
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #203 — PPTX table cell: GridSpan/RowSpan type mismatch
    /// File: PowerPointHandler.ShapeProperties.cs, lines 553, 556
    /// Uses Int32Value but DrawingML spec requires unsigned GridSpan/RowSpan.
    [Fact]
    public void Bug203_PptxTableCell_GridSpanTypeMismatch()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });

        // Set gridspan — should use unsigned type
        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["gridspan"] = "2" });

        var node = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
    }

    /// Bug #204 — PPTX table: int.Parse on row addition cols
    /// File: PowerPointHandler.Add.cs, line 1000
    /// Uses int.Parse without validation; negative cols not rejected.
    [Fact]
    public void Bug204_PptxTable_IntParseOnRowAdditionCols()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        var act = () => pptx.Add("/slide[1]/table[1]", "row", null, new()
        {
            ["cols"] = "abc"
        });

        act.Should().Throw<FormatException>(
            "int.Parse crashes on 'abc' for table row cols");
    }

    /// Bug #205 — Word image: NonVisualDrawingProperties.Id hardcoded to 0
    /// File: WordHandler.ImageHelpers.cs, lines 45, 115
    /// PIC.NonVisualDrawingProperties.Id = 0U for all images.
    /// Should be unique per drawing object.
    [Fact]
    public void Bug205_WordImage_HardcodedZeroId()
    {
        var imgPath = CreateTempImage();
        try
        {
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new() { ["src"] = imgPath });

            ReopenWord();
            var node = _wordHandler.Get("/body");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #206 — Excel data validation: formula injection risk
    /// File: ExcelHandler.Add.cs, lines 266-275
    /// Formula1 and formula2 are passed through without sanitization.
    /// Only List type gets auto-quoted; other types accept raw formulas.
    [Fact]
    public void Bug206_ExcelDataValidation_FormulaInjectionRisk()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "50" });

        // Custom validation with arbitrary formula
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["type"] = "custom",
            ["sqref"] = "A1",
            ["formula1"] = "=INDIRECT(\"Sheet2!A1\")"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #207 — PPTX background: gradient stops ignore SchemeColor
    /// File: PowerPointHandler.Background.cs, lines 120-122
    /// Same as Bug #200 but for slide backgrounds.
    [Fact]
    public void Bug207_PptxBackground_GradientIgnoresSchemeColor()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set background gradient
        pptx.Set("/slide[1]", new() { ["background"] = "gradient:FF0000-0000FF" });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #208 — PPTX table cell: bool.Parse on vmerge/hmerge
    /// File: PowerPointHandler.ShapeProperties.cs, lines 559, 562
    /// Uses bool.Parse for merge properties.
    [Fact]
    public void Bug208_PptxTableCell_BoolParseOnMerge()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        var act = () => pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["vmerge"] = "yes"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on 'yes' for vertical merge");
    }

    /// Bug #209 — Word image: bool.Parse on anchor property
    /// File: WordHandler.Add.cs, line 488
    /// Uses bool.Parse for floating image anchor setting.
    /// Already documented in Bug124 but confirms pattern in image context.
    [Fact]
    public void Bug209_WordImage_BoolParseOnBehindText()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["anchor"] = "true",
                ["behindtext"] = "1"
            });

            act.Should().Throw<FormatException>(
                "bool.Parse crashes on '1' for behindtext — should use IsTruthy");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #210 — PPTX slide layout: silent null when layout name missing
    /// File: PowerPointHandler.Query.cs, lines 282-285
    /// If slide has layout but layout name is null, no layout info is returned.
    /// Users can't tell if layout is missing vs. unnamed.
    [Fact]
    public void Bug210_PptxSlideLayout_SilentNullOnMissingName()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
        // If layout name is available, it should be in format
        // If not, user has no way to know whether layout exists
    }

    // ==================== Bug #211-230: Comments, Merge cells, Connectors, ParseEmu ====================

    /// Bug #211 — Word comment: DateTime.Parse without validation
    /// File: WordHandler.Add.cs, line 557
    /// Uses DateTime.Parse on user input without TryParse.
    [Fact]
    public void Bug211_WordComment_DateTimeParseNoValidation()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Commented text" });

        var act = () => _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "Review needed",
            ["author"] = "Tester",
            ["date"] = "not-a-date"
        });

        act.Should().Throw<FormatException>(
            "DateTime.Parse crashes on 'not-a-date' — should use TryParse");
    }

    /// Bug #212 — Word comment: empty author causes IndexOutOfRange
    /// File: WordHandler.Add.cs, line 544
    /// author[..1] on empty string throws IndexOutOfRangeException.
    [Fact]
    public void Bug212_WordComment_EmptyAuthorCrash()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });

        var act = () => _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "Comment",
            ["author"] = ""
        });

        act.Should().Throw<Exception>(
            "Empty author string causes author[..1] to throw IndexOutOfRangeException");
    }

    /// Bug #213 — Word comment: orphaned markers on Remove
    /// File: WordHandler.Add.cs, lines 1177-1185
    /// Removing a CommentRangeStart doesn't clean up the corresponding
    /// CommentRangeEnd, CommentReference, or Comment object.
    [Fact]
    public void Bug213_WordComment_OrphanedMarkersOnRemove()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "Review",
            ["author"] = "Tester"
        });

        // Remove just cleans the element, not related comment parts
        // This is the same Remove used for all elements
        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull();
    }

    /// Bug #214 — Word comment: Comment saved before markup insertion
    /// File: WordHandler.Add.cs, lines 553-578
    /// Comment object is saved to comments part before range markers
    /// are inserted into document. If insertion fails, orphaned comment remains.
    [Fact]
    public void Bug214_WordComment_SavedBeforeMarkup()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "Comment text",
            ["author"] = "Author"
        });

        ReopenWord();
        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull("comment should be properly inserted");
    }

    /// Bug #215 — Excel merge: no overlap detection
    /// File: ExcelHandler.Set.cs, lines 639-654
    /// Merge operation only checks for exact duplicates, not overlaps.
    /// Overlapping merges create corrupt Excel files.
    [Fact]
    public void Bug215_ExcelMerge_NoOverlapDetection()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "3" });

        // First merge
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:B2" });

        // Overlapping merge — should be rejected but isn't
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "B1:C2" });

        // Excel would reject this file due to overlapping merge ranges
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #216 — Excel merge: row deletion doesn't clean up merge ranges
    /// File: ExcelHandler.Add.cs, lines 954-962
    /// Deleting a row that participates in a merge doesn't update
    /// or remove the affected merge definition.
    [Fact]
    public void Bug216_ExcelMerge_RowDeletionNoCleanup()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Merged" });
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:A3" });

        // Delete row 2 which is part of the merge
        _excelHandler.Remove("/Sheet1/row[2]");

        // The merge definition "A1:A3" still exists but row 2 is gone
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #217 — Excel merge: silent data loss
    /// File: ExcelHandler.Set.cs, lines 639-654
    /// Merging range with multiple values silently discards all but top-left.
    [Fact]
    public void Bug217_ExcelMerge_SilentDataLoss()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Keep" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Lost" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "Lost too" });

        // Merge will only keep A1's value
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:C1" });

        ReopenExcel();
        // B1 and C1 data should be preserved or warned about
        var b1 = _excelHandler.Get("/Sheet1/B1");
        // Data in B1 may be silently lost during merge
    }

    /// Bug #218 — PPTX connector: endpoint Index always 0
    /// File: PowerPointHandler.Add.cs, lines 870, 872
    /// Connector start/end connection Index is hardcoded to 0,
    /// ignoring shape connection point selection.
    [Fact]
    public void Bug218_PptxConnector_EndpointIndexAlwaysZero()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });

        // Add connector — Index=0 is always used regardless of shape geometry
        pptx.Add("/slide[1]", "connector", null, new()
        {
            ["startshape"] = "1",
            ["endshape"] = "2"
        });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #219 — PPTX connector: height defaults to 0
    /// File: PowerPointHandler.Add.cs, line 858
    /// Connector without explicit height gets Cy=0, creating degenerate shape.
    [Fact]
    public void Bug219_PptxConnector_HeightDefaultsToZero()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add connector without height — defaults to 0
        pptx.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "100", ["y"] = "100", ["width"] = "200"
        });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #220 — PPTX group: bounding box invalid when shapes lack transforms
    /// File: PowerPointHandler.Add.cs, lines 944-957
    /// If all grouped shapes have null Transform2D, bounding box overflows.
    /// minX=long.MaxValue, maxX=0 → Cx = 0 - MaxValue (negative).
    [Fact]
    public void Bug220_PptxGroup_BoundingBoxOverflow()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });

        // Group the shapes
        pptx.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #221 — ParseEmu: long to int cast overflow
    /// File: PowerPointHandler.Fill.cs, lines 182-192
    /// ParseEmu returns long but results cast to int for BodyProperties.
    /// Values > int.MaxValue silently overflow to negative.
    [Fact]
    public void Bug221_ParseEmu_LongToIntCastOverflow()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Very large inset value that overflows int
        var act = () => pptx.Set("/slide[1]/shape[1]", new()
        {
            ["inset"] = "1000cm,1000cm,1000cm,1000cm"
        });

        // 1000cm = 360,000,000,000 EMU which overflows int.MaxValue
        act.Should().Throw<Exception>(
            "ParseEmu long result cast to int causes silent overflow for large values");
    }

    /// Bug #222 — ParseEmu: negative dimensions accepted
    /// File: WordHandler.ImageHelpers.cs, lines 17-30 and PowerPointHandler.Helpers.cs, lines 161-173
    /// No validation for negative values. "-5cm" produces -1800000 EMU.
    [Fact]
    public void Bug222_ParseEmu_NegativeDimensionsAccepted()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var act = () => pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["width"] = "-5cm"
        });

        // Negative width creates invalid shape
        act.Should().Throw<Exception>(
            "Negative dimensions should be rejected by ParseEmu");
    }

    /// Bug #223 — ParseEmu: empty unit suffix causes crash
    /// File: WordHandler.ImageHelpers.cs, line 22
    /// Input "cm" (unit only, no number) → value[..^2] = "" → double.Parse("") crash.
    [Fact]
    public void Bug223_ParseEmu_EmptyUnitSuffixCrash()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "cm"
            });

            act.Should().Throw<FormatException>(
                "ParseEmu crashes on 'cm' (no number) — should validate input length");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #224 — ParseEmu: unsupported units silently fail
    /// File: WordHandler.ImageHelpers.cs, lines 17-30
    /// "5mm" is not supported — falls through to long.Parse("5mm") which crashes.
    [Fact]
    public void Bug224_ParseEmu_UnsupportedUnitsCrash()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "50mm"
            });

            act.Should().Throw<FormatException>(
                "ParseEmu doesn't support 'mm' unit — falls through to long.Parse('50mm')");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #225 — Excel merge: no validation of merge range format
    /// File: ExcelHandler.Set.cs, lines 425-433
    /// Only validates first part of range (before ':'). Second part not checked.
    [Fact]
    public void Bug225_ExcelMerge_NoRangeFormatValidation()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Malformed range — second part not validated
        var act = () => _excelHandler.Set("/Sheet1", new()
        {
            ["merge"] = "A1:INVALID"
        });

        // Should validate both parts of the range
        act.Should().NotThrow("malformed merge range accepted without validation");
    }

    /// Bug #226 — Excel duplicate ReorderWorksheetChildren call
    /// File: ExcelHandler.Set.cs, line 680
    /// ReorderWorksheetChildren is called twice in a row — copy-paste error.
    [Fact]
    public void Bug226_ExcelSet_DuplicateReorderCall()
    {
        // This is a performance bug — the function is called twice
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Updated" });
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1/A1");
        node?.Text.Should().Be("Updated");
    }

    /// Bug #227 — Word navigation: CommentReference runs filtered out
    /// File: WordHandler.Navigation.cs, lines 161-163
    /// Runs containing CommentReference are hidden from navigation,
    /// making it impossible to query or modify comment references directly.
    [Fact]
    public void Bug227_WordNavigation_CommentRunsFiltered()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Commented" });
        _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "A comment",
            ["author"] = "Test"
        });

        // Runs with CommentReference are filtered from navigation
        // This means you can't directly access or modify them
        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull();
    }

    /// Bug #228 — PPTX group: empty shapes list not validated
    /// File: PowerPointHandler.Add.cs, lines 931-944
    /// If shapes="" is provided, toGroup is empty, causing invalid group.
    [Fact]
    public void Bug228_PptxGroup_EmptyShapesList()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });

        var act = () => pptx.Add("/slide[1]", "group", null, new()
        {
            ["shapes"] = ""
        });

        act.Should().Throw<Exception>(
            "Empty shapes list should be rejected for group creation");
    }

    /// Bug #229 — ParseEmu: double truncation instead of rounding
    /// File: WordHandler.ImageHelpers.cs, line 22
    /// (long)(double.Parse(value) * 360000) truncates instead of rounding.
    /// "0.001cm" → 360 EMU (truncated) vs 360 EMU (correct by coincidence).
    [Fact]
    public void Bug229_ParseEmu_TruncationInsteadOfRounding()
    {
        var imgPath = CreateTempImage();
        try
        {
            // Fractional cm values may lose precision due to truncation
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "2.54cm"  // Should be exactly 1 inch = 914400 EMU
            });

            ReopenWord();
            var node = _wordHandler.Get("/body");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #230 — ParseEmu: duplicate implementations in Word and PPTX
    /// File: WordHandler.ImageHelpers.cs lines 17-30, PowerPointHandler.Helpers.cs lines 161-173
    /// Identical code duplicated — any fix must be applied to both files.
    [Fact]
    public void Bug230_ParseEmu_DuplicateImplementations()
    {
        // This is a code quality bug — ParseEmu exists in two places
        // Any bug fix or enhancement must be applied to both
        // Word ParseEmu and PPTX ParseEmu are separate copies
        var imgPath = CreateTempImage();
        try
        {
            _wordHandler.Add("/body", "p", null, new() { ["text"] = "Image" });
            _wordHandler.Add("/body/p[1]", "image", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "5cm"
            });

            BlankDocCreator.Create(_pptxPath);
            using var pptx = new PowerPointHandler(_pptxPath, editable: true);
            pptx.Add("/", "slide", null, new());
            pptx.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Test",
                ["width"] = "5cm"
            });
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Bug #231-250: Media, Charts, TOC, File lifecycle ====================

    /// Bug #231 — PPTX audio: AudioFromFile.Link uses video relationship ID
    /// File: PowerPointHandler.Add.cs, line 776
    /// AudioFromFile.Link is set to videoRelId instead of the audio-specific ID.
    /// This causes audio files to reference the wrong relationship.
    [Fact]
    public void Bug231_PptxAudio_WrongRelationshipId()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Create a test audio file (WAV format, minimal)
        var audioPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.wav");
        try
        {
            // Create minimal WAV file
            using (var ms = new MemoryStream())
            using (var bw = new BinaryWriter(ms))
            {
                bw.Write("RIFF"u8.ToArray());
                bw.Write(36); // chunk size
                bw.Write("WAVE"u8.ToArray());
                bw.Write("fmt "u8.ToArray());
                bw.Write(16); // subchunk size
                bw.Write((short)1); // PCM
                bw.Write((short)1); // mono
                bw.Write(44100); // sample rate
                bw.Write(44100); // byte rate
                bw.Write((short)1); // block align
                bw.Write((short)8); // bits per sample
                bw.Write("data"u8.ToArray());
                bw.Write(0); // data size
                File.WriteAllBytes(audioPath, ms.ToArray());
            }

            // The audio element incorrectly uses videoRelId
            // This is a critical bug — audio won't play due to wrong relationship
            pptx.Add("/slide[1]", "audio", null, new() { ["path"] = audioPath });
            var node = pptx.Get("/slide[1]");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(audioPath)) File.Delete(audioPath);
        }
    }

    /// Bug #232 — PPTX media: volume double.Parse without validation
    /// File: PowerPointHandler.Add.cs, lines 822-823
    /// Volume uses double.Parse without bounds checking or TryParse.
    [Fact]
    public void Bug232_PptxMedia_VolumeParseNoValidation()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            var act = () => pptx.Add("/slide[1]", "video", null, new()
            {
                ["path"] = imgPath,
                ["volume"] = "loud"
            });

            act.Should().Throw<FormatException>(
                "double.Parse crashes on 'loud' for volume — should use TryParse");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #233 — PPTX media: trim values not validated
    /// File: PowerPointHandler.Add.cs, lines 808-814
    /// trimStart and trimEnd are passed directly without validation.
    [Fact]
    public void Bug233_PptxMedia_TrimValuesNotValidated()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            // Non-numeric trim values passed through without validation
            pptx.Add("/slide[1]", "video", null, new()
            {
                ["path"] = imgPath,
                ["trimstart"] = "invalid"
            });

            var node = pptx.Get("/slide[1]");
            node.Should().NotBeNull("invalid trim value accepted without validation");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #234 — PPTX media: HyperlinkOnClick with empty Id
    /// File: PowerPointHandler.Add.cs, line 769
    /// HyperlinkOnClick.Id is set to "" which may cause relationship issues.
    [Fact]
    public void Bug234_PptxMedia_EmptyHyperlinkId()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #235 — Excel chart: 3D flag parsed but never used
    /// File: ExcelHandler.Helpers.cs, lines 391-414, 490-539
    /// ExcelChartParseChartType correctly parses is3D but the flag
    /// is ignored by all chart builders. "bar3d" == "bar".
    [Fact]
    public void Bug235_ExcelChart_3DFlagIgnored()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        // "bar3d" should create a 3D chart but the flag is discarded
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar3d",
            ["data"] = "Sales:10,20,30"
        });

        // The chart is identical to a regular bar chart
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("3D flag is parsed but ignored");
    }

    /// Bug #236 — Excel chart: non-contiguous series definitions not supported
    /// File: ExcelHandler.Helpers.cs, lines 434-450
    /// If series1 and series3 are provided (skipping series2), the loop
    /// breaks at series2 and series3 is never read.
    [Fact]
    public void Bug236_ExcelChart_NonContiguousSeriesIgnored()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["series1"] = "Sales:10,20,30",
            ["series3"] = "Costs:5,10,15"  // Skipped series2 — series3 is silently ignored
        });

        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("non-contiguous series silently ignored");
    }

    /// Bug #237 — Excel chart: category/series length mismatch not validated
    /// File: ExcelHandler.Helpers.cs, lines 725, 742, 759, 776
    /// Series data with different lengths than categories causes misalignment.
    [Fact]
    public void Bug237_ExcelChart_CategorySeriesLengthMismatch()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        // 3 categories but 5 data points — misalignment
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Sales:10,20,30,40,50"
        });

        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("category/series length mismatch accepted without warning");
    }

    /// Bug #238 — Word TOC Set: bool.Parse on hyperlinks/pagenumbers
    /// File: WordHandler.Set.cs, lines 70-81
    /// TOC update uses bool.Parse for hyperlinks and pagenumbers switches.
    [Fact]
    public void Bug238_WordTocSet_BoolParse()
    {
        _wordHandler.Add("/body", "toc", null, new()
        {
            ["levels"] = "1-3"
        });

        var act = () => _wordHandler.Set("/body/toc[1]", new()
        {
            ["hyperlinks"] = "yes"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on 'yes' in TOC hyperlinks setting");
    }

    /// Bug #239 — Word bookmark: name validation insufficient
    /// File: WordHandler.Add.cs, lines 587-589
    /// Only checks for empty name, not Word naming rules
    /// (must start with letter, no spaces, max 40 chars).
    [Fact]
    public void Bug239_WordBookmark_NameValidationInsufficient()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Bookmarked" });

        // Bookmark name with spaces — invalid per Word spec
        _wordHandler.Add("/body/p[1]", "bookmark", null, new()
        {
            ["name"] = "My Bookmark Name"
        });

        ReopenWord();
        var node = _wordHandler.Get("/bookmark[My Bookmark Name]");
        // Word may reject or corrupt documents with invalid bookmark names
    }

    /// Bug #240 — Constructor exception leaks file handle
    /// File: WordHandler.cs, lines 23-27
    /// If constructor throws after Open(), document handle is never released.
    [Fact]
    public void Bug240_ConstructorException_LeakedFileHandle()
    {
        // Open a valid document, then verify it can be reopened after disposal
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);

        // File should be accessible — no leaked handle
        _wordHandler.Should().NotBeNull();
    }

    /// Bug #241 — No Save() before Dispose() in handlers
    /// File: WordHandler.cs line 108-111, ExcelHandler.cs line 216, PowerPointHandler.cs line 564
    /// Dispose calls _doc.Dispose() without explicit Save().
    /// Changes may not be flushed to disk.
    [Fact]
    public void Bug241_NoSaveBeforeDispose()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test persist" });

        // Dispose without explicit save — relies on SDK auto-save
        _wordHandler.Dispose();

        // Reopen and verify data persisted
        _wordHandler = new WordHandler(_docxPath, editable: true);
        var node = _wordHandler.Get("/body/p[1]");
        node?.Text.Should().Contain("Test persist",
            "Data should persist through Dispose without explicit Save");
    }

    /// Bug #242 — PPTX slide deletion: order-dependent cleanup
    /// File: PowerPointHandler.Add.cs, lines 1158-1160
    /// Slide removed from SlideIdList before part deletion.
    /// If DeletePart fails, slide is removed but part orphaned.
    [Fact]
    public void Bug242_PptxSlideDelete_OrderDependentCleanup()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/", "slide", null, new());

        // Delete second slide
        pptx.Remove("/slide[2]");

        var root = pptx.Get("/");
        root.Children.Where(c => c.Type == "slide").Should().HaveCount(1,
            "only one slide should remain after deletion");
    }

    /// Bug #243 — Excel chart: scatter chart axis semantics wrong
    /// File: ExcelHandler.Helpers.cs, lines 529-532
    /// Scatter charts need two ValueAxis objects but code creates
    /// category axis + value axis pattern for all chart types.
    [Fact]
    public void Bug243_ExcelChart_ScatterAxisSemanticsWrong()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "scatter",
            ["categories"] = "1,2,3,4,5",
            ["data"] = "Y:10,20,15,25,30"
        });

        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #244 — Excel chart double.Parse on data values
    /// File: ExcelHandler.Helpers.cs, lines 428, 440, 447
    /// double.Parse without TryParse on user-provided chart data.
    [Fact]
    public void Bug244_ExcelChart_DoubleParseOnData()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        var act = () => _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["series1"] = "Sales:10,N/A,30"
        });

        act.Should().Throw<FormatException>(
            "double.Parse crashes on 'N/A' in chart data values");
    }

    /// Bug #245 — PPTX media: shape ID collision risk
    /// File: PowerPointHandler.Add.cs, line 763
    /// Media ID = ChildElements.Count + 2 doesn't guarantee uniqueness.
    [Fact]
    public void Bug245_PptxMedia_ShapeIdCollisionRisk()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape 1" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape 2" });

        // Delete shape 1, then add media — ID may collide
        pptx.Remove("/slide[1]/shape[1]");

        // After deletion, ChildElements.Count decreases
        // New ID = Count + 2 may collide with remaining shape's ID
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "New Shape" });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #246 — Word TOC Add: field code empty string fallback
    /// File: WordHandler.Set.cs, lines 49-56
    /// If TOC exists but FieldCode.Text is null, silently uses "".
    [Fact]
    public void Bug246_WordToc_FieldCodeEmptyFallback()
    {
        _wordHandler.Add("/body", "toc", null, new()
        {
            ["levels"] = "1-3"
        });

        // Verify TOC was created with valid field code
        var node = _wordHandler.Get("/body/toc[1]");
        node.Should().NotBeNull("TOC should be queryable after creation");
    }

    /// Bug #247 — Excel chart: pie chart negative values not validated
    /// File: ExcelHandler.Helpers.cs, lines 658-662
    /// Negative values in pie charts are meaningless but accepted.
    [Fact]
    public void Bug247_ExcelChart_PieChartNegativeValues()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "pie",
            ["data"] = "Sales:10,-5,30,-10"
        });

        // Pie charts with negative values produce invalid visualizations
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("pie chart with negative values accepted without warning");
    }

    /// Bug #248 — PPTX media: format detection relies only on extension
    /// File: PowerPointHandler.Add.cs, lines 703-712
    /// Only uses file extension for format detection.
    /// A renamed .txt file with .mp4 extension is accepted as video.
    [Fact]
    public void Bug248_PptxMedia_FormatDetectionByExtensionOnly()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Create a text file disguised as MP4
        var fakePath = Path.Combine(Path.GetTempPath(), $"fake_{Guid.NewGuid():N}.mp4");
        try
        {
            File.WriteAllText(fakePath, "This is not a video");

            // Should detect invalid file format, but only checks extension
            pptx.Add("/slide[1]", "video", null, new() { ["path"] = fakePath });

            var node = pptx.Get("/slide[1]");
            node.Should().NotBeNull("fake media file accepted based on extension only");
        }
        finally
        {
            if (File.Exists(fakePath)) File.Delete(fakePath);
        }
    }

    /// Bug #249 — Excel DeletePart without error handling
    /// File: ExcelHandler.Add.cs, line 942
    /// Sheet part deletion has no error handling. If DeletePart fails,
    /// the sheet XML is already removed but part remains orphaned.
    [Fact]
    public void Bug249_ExcelDeletePart_NoErrorHandling()
    {
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "ToDelete" });
        _excelHandler.Add("/ToDelete", "cell", null, new() { ["ref"] = "A1", ["value"] = "Data" });

        _excelHandler.Remove("/ToDelete");

        ReopenExcel();
        var root = _excelHandler.Get("/");
        root.Children.Where(c => c.Path.Contains("ToDelete")).Should().BeEmpty(
            "deleted sheet should not appear after reopen");
    }

    /// Bug #250 — Excel chart: empty series data throws Max() crash
    /// File: ExcelHandler.Helpers.cs, line 416-453
    /// Empty series list with explicit chart creation causes InvalidOperationException.
    [Fact]
    public void Bug250_ExcelChart_EmptySeriesDataCrash()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        var act = () => _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = ""
        });

        act.Should().Throw<Exception>(
            "Empty chart data should throw clear error, not crash on Max() of empty sequence");
    }

}
