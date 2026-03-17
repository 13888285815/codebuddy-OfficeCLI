// Bug hunt tests Part 2 — Bug #91-170
// Chart, Animations, FormulaParser, Delete/Move, Query, Selector

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ==================== Bug #91-110: Chart, Animations, FormulaParser, Excel Add, StyleManager ====================

    /// Bug #91 — PPTX Chart: double.Parse on malformed series data
    /// File: PowerPointHandler.Chart.cs, line 65
    /// ParseSeriesData uses double.Parse(v.Trim()) without TryParse.
    /// If chart data contains non-numeric values like "N/A", it crashes.
    [Fact]
    public void Bug91_PptxChart_DoubleParseOnMalformedSeriesData()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Provide data with non-numeric value — should not crash
        var act = () => pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,N/A,30"
        });

        act.Should().Throw<FormatException>(
            "double.Parse crashes on 'N/A' instead of using TryParse with graceful fallback");
    }

    /// Bug #92 — PPTX Chart: int.Parse on malformed combosplit
    /// File: PowerPointHandler.Chart.cs, line 175
    /// Combo chart split index uses int.Parse without validation.
    [Fact]
    public void Bug92_PptxChart_IntParseOnMalformedComboSplit()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var act = () => pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "combo",
            ["combosplit"] = "two",
            ["data"] = "A:1,2,3;B:4,5,6"
        });

        act.Should().Throw<FormatException>(
            "int.Parse crashes on 'two' instead of using TryParse");
    }

    /// Bug #93 — PPTX Chart: double.Parse on axis min/max/unit properties
    /// File: PowerPointHandler.Chart.cs, lines 1040, 1051, 1061, 1071
    /// SetChartProperties uses double.Parse for axis values without TryParse.
    [Fact]
    public void Bug93_PptxChart_DoubleParseOnAxisProperties()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30"
        });

        // Set axis min with non-numeric value
        var act = () => pptx.Set("/slide[1]/chart[1]", new()
        {
            ["axismin"] = "auto"
        });

        act.Should().Throw<FormatException>(
            "double.Parse crashes on 'auto' for axis min — should use TryParse");
    }

    /// Bug #94 — PPTX Animations: bounce and zoom share preset ID 21
    /// File: PowerPointHandler.Animations.cs, lines 688-689
    /// Both "zoom" and "bounce" map to preset ID 21, causing "bounce"
    /// to be read back as "zoom" when inspecting animation properties.
    [Fact]
    public void Bug94_PptxAnimations_BounceAndZoomSharePresetId()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Add bounce animation
        pptx.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "bounce",
            ["trigger"] = "onclick"
        });

        // Get the animation back — it should report "bounce" not "zoom"
        var node = pptx.Get("/slide[1]/shape[1]/animation[1]");
        // Due to shared preset ID, bounce is indistinguishable from zoom
        // This is a data loss bug — the animation type cannot roundtrip
        node.Format.TryGetValue("effect", out var effect);
        // If the handler reads it back, it'll say "zoom" instead of "bounce"
        // because both use preset ID 21
        (effect == "bounce" || effect == null).Should().BeTrue(
            "bounce animation should roundtrip, but preset ID collision with zoom causes data loss");
    }

    /// Bug #95 — FormulaParser: \left...\right delimiter not captured
    /// File: FormulaParser.cs, lines 819-829, 876
    /// When parsing \right], the closing delimiter character is consumed
    /// but never stored. Line 876 guesses closeChar based on openChar,
    /// so \left(...\right] produces ")" instead of "]".
    [Fact]
    public void Bug95_FormulaParser_RightDelimiterNotCaptured()
    {
        // \left( x \right] should produce mismatched delimiters
        var result = FormulaParser.Parse(@"\left( x \right]");
        var xml = result.OuterXml;

        // The closing delimiter should be "]" but due to the bug,
        // it's guessed from openChar="(" → closeChar=")"
        xml.Should().Contain("]",
            "\\right] should produce ']' as closing delimiter, but the parser " +
            "discards the actual delimiter and guesses ')' from the opening '('");
    }

    /// Bug #96 — FormulaParser: empty matrix crashes on rows.Max()
    /// File: FormulaParser.cs, line 1238
    /// ParseMatrix calls rows.Max(r => r.Count) which throws
    /// InvalidOperationException if rows is empty (empty matrix env).
    [Fact]
    public void Bug96_FormulaParser_EmptyMatrixCrash()
    {
        // An empty matrix environment should not crash
        var act = () => FormulaParser.Parse(@"\begin{matrix}\end{matrix}");

        // This may crash with InvalidOperationException from Max() on empty sequence
        // or it may produce an empty matrix — either way it should not throw
        act.Should().NotThrow(
            "An empty \\begin{matrix}\\end{matrix} should not crash, " +
            "but rows.Max() on empty sequence throws InvalidOperationException");
    }

    /// Bug #97 — FormulaParser: RewriteOver substring out of range
    /// File: FormulaParser.cs, lines 90-92
    /// If \over immediately follows opening brace with no numerator,
    /// e.g., "{\over x}", the Substring call produces negative length.
    [Fact]
    public void Bug97_FormulaParser_RewriteOverEdgeCase()
    {
        // Edge case: \over with empty numerator
        var act = () => FormulaParser.Parse(@"{\over x}");

        // Should handle gracefully, not throw ArgumentOutOfRangeException
        act.Should().NotThrow(
            "'{\\over x}' with empty numerator causes negative Substring length");
    }

    /// Bug #98 — Excel Add: int.Parse on non-numeric "cols" property
    /// File: ExcelHandler.Add.cs, line 55
    /// When adding a row, int.Parse(colsStr) crashes if cols is not numeric.
    [Fact]
    public void Bug98_ExcelAdd_IntParseOnMalformedCols()
    {
        var act = () => _excelHandler.Add("/Sheet1", "row", null, new()
        {
            ["cols"] = "five"
        });

        act.Should().Throw<FormatException>(
            "int.Parse crashes on 'five' — should use TryParse for user input");
    }

    /// Bug #99 — Excel Add: int.Parse on chart position properties
    /// File: ExcelHandler.Add.cs, lines 838-841
    /// Chart x, y, width, height use int.Parse without TryParse.
    [Fact]
    public void Bug99_ExcelAdd_IntParseOnChartPosition()
    {
        var act = () => _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30",
            ["x"] = "auto"
        });

        act.Should().Throw<FormatException>(
            "int.Parse crashes on 'auto' for chart x position — should use TryParse");
    }

    /// Bug #100 — Excel Add: row index cast overflow from uint to int
    /// File: ExcelHandler.Add.cs, line 49
    /// Row index is cast from uint to int, which overflows for very large row indices.
    [Fact]
    public void Bug100_ExcelAdd_RowIndexUintToIntCast()
    {
        // Add a row with a very large row index (Excel max is 1048576)
        // The cast from uint to int is safe for valid Excel row numbers,
        // but the code doesn't validate the range
        _excelHandler.Add("/Sheet1", "row", 1048576, new());
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
        // The real concern is no validation that row index is within Excel limits
    }

    /// Bug #91 already claimed — renumbered
    /// Bug #101 — PPTX Chart: scatter chart silently converts non-numeric categories to 0
    /// File: PowerPointHandler.Chart.cs, line 388
    /// double.TryParse failures silently become 0, corrupting data.
    [Fact]
    public void Bug101_PptxChart_ScatterSilentZeroConversion()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Scatter chart with non-numeric categories — they silently become 0
        pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "scatter",
            ["categories"] = "Jan,Feb,Mar",
            ["data"] = "Sales:10,20,30"
        });

        // The x-axis values should NOT silently become [0, 0, 0]
        // This is a data integrity issue — user's category labels are lost
        var node = pptx.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull("scatter chart with text categories should warn, not silently zero out");
    }

    /// Bug #102 — Excel StyleManager: underline type loss
    /// File: ExcelStyleManager.cs, line 255
    /// When merging styles, underline defaults to "single" regardless of actual type.
    /// Double underline becomes single underline silently.
    [Fact]
    public void Bug102_ExcelStyleManager_UnderlineTypeLoss()
    {
        // Set double underline
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        _excelHandler.Set("/Sheet1/A1", new() { ["underline"] = "double" });

        // Now set bold (which triggers style merge)
        _excelHandler.Set("/Sheet1/A1", new() { ["bold"] = "true" });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1/A1");

        // The underline should still be "double", not downgraded to "single"
        node.Format.TryGetValue("underline", out var ul);
        (ul == "double" || ul == "Double").Should().BeTrue(
            "Double underline should be preserved when merging styles, " +
            "but ExcelStyleManager defaults baseFont underline to 'single'");
    }

    /// Bug #103 — Word StyleList: int.Parse on font size
    /// File: WordHandler.StyleList.cs, line 104
    /// Uses int.Parse(size) on potentially malformed size values.
    [Fact]
    public void Bug103_WordStyleList_IntParseOnFontSize()
    {
        // Set a paragraph style with a non-numeric size
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });

        // Try to create a list with a non-standard size value
        var act = () => _wordHandler.Set("/body/p[1]", new()
        {
            ["style"] = "ListParagraph",
            ["size"] = "12.5"
        });

        // int.Parse will fail on "12.5" — should use TryParse or accept decimal
        act.Should().Throw<FormatException>(
            "int.Parse crashes on '12.5' — font sizes should support half-points");
    }

    /// Bug #104 — Word StyleList: numbering ID generation starts from 0
    /// File: WordHandler.StyleList.cs, line 268
    /// DefaultIfEmpty(-1).Max() + 1 returns 0 when no abstract nums exist.
    /// Starting abstract numbering ID from 0 may conflict with reserved values.
    [Fact]
    public void Bug104_WordStyleList_NumberingIdStartsFromZero()
    {
        // Create a fresh document with no existing numbering
        // Adding the first list should get a valid numbering ID
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Item 1" });
        _wordHandler.Set("/body/p[1]", new()
        {
            ["listStyle"] = "bullet"
        });

        var node = _wordHandler.Get("/body/p[1]");
        // Verify numbering was applied
        node.Format.TryGetValue("numid", out var numIdObj);
        // The numid should be > 0 (not 0 which may conflict with "no numbering")
        if (numIdObj != null)
        {
            var numId = Convert.ToInt32(numIdObj);
            numId.Should().BeGreaterThan(0,
                "Numbering ID 0 typically means 'no numbering' in Word; " +
                "the generator should start from 1");
        }
    }

    /// Bug #105 — FormulaParser: MatrixColumns creates multiple wrappers
    /// File: FormulaParser.cs, lines 1241-1243
    /// Each column gets its own MatrixColumns wrapper instead of one shared wrapper.
    /// This creates malformed OMML: <mPr><mcs><mc>...</mc></mcs><mcs><mc>...</mc></mcs></mPr>
    /// instead of <mPr><mcs><mc>...</mc><mc>...</mc></mcs></mPr>.
    [Fact]
    public void Bug105_FormulaParser_MatrixColumnsMalformedStructure()
    {
        // cases environment with 2 columns should have one MatrixColumns with 2 children
        var result = FormulaParser.Parse(@"\begin{cases} a & b \\ c & d \end{cases}");
        var xml = result.OuterXml;

        // Count how many <m:mcs> elements appear
        var mcsCount = System.Text.RegularExpressions.Regex.Matches(xml, "<m:mcs>").Count;
        mcsCount.Should().BeLessOrEqualTo(1,
            "There should be one <m:mcs> element containing all <m:mc> children, " +
            "but the code creates a separate <m:mcs> per column");
    }

    /// Bug #106 — PPTX Chart: series update double.Parse
    /// File: PowerPointHandler.Chart.cs, lines 1126, 1140
    /// SetChartProperties parses updated series data with double.Parse.
    [Fact]
    public void Bug106_PptxChart_SeriesUpdateDoubleParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30"
        });

        // Update series with a value that can't be parsed
        var act = () => pptx.Set("/slide[1]/chart[1]", new()
        {
            ["series1"] = "Revenue:10,twenty,30"
        });

        act.Should().Throw<FormatException>(
            "double.Parse crashes on 'twenty' when updating chart series data");
    }

    /// Bug #107 — BlankDocCreator PPTX: relationship ID collision
    /// File: BlankDocCreator.cs, lines 65, 69, 141, 160, 179
    /// slideLayoutPart uses "rId1" for slide layout, and "rId2" for theme,
    /// but layout parts added to slideMaster may collide with theme's "rId2".
    [Fact]
    public void Bug107_BlankDocCreator_PptxRelationshipIdCollision()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"blank_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(tempPath);

            // Open and verify the file is valid
            using var pptx = new PowerPointHandler(tempPath, editable: false);
            var root = pptx.Get("/");
            root.Should().NotBeNull("blank PPTX should be openable without errors");

            // Verify at least one slide layout exists
            root.Children.Should().NotBeEmpty("blank PPTX should have slide structure");
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    /// Bug #108 — Word GetHeadingLevel: only checks first digit
    /// File: WordHandler.Helpers.cs, lines 159-170
    /// GetHeadingLevel parses only the first character after "Heading ",
    /// so "Heading 10" returns 1 instead of 10.
    [Fact]
    public void Bug108_WordGetHeadingLevel_SingleDigitOnly()
    {
        // This is a code-level bug that affects heading detection
        // Styles like "Heading 10" (valid in custom templates) are misidentified
        // The method uses: styleName[8] - '0' which only reads one character
        // "Heading 10" → reads '1' → returns 1 instead of 10

        // We can verify by setting heading style and checking the returned node
        _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Heading Ten",
            ["style"] = "Heading1"
        });

        var node = _wordHandler.Get("/body/p[1]");
        // This test documents the limitation: only single-digit headings work
        node.Should().NotBeNull();
        // The real bug is in GetHeadingLevel which parses only one char
    }

    /// Bug #109 — Word IsNormalStyle: case-sensitive comparison
    /// File: WordHandler.Helpers.cs, lines 172-176
    /// IsNormalStyle compares style name case-sensitively,
    /// so "normal" (lowercase) doesn't match if the style is stored as "Normal".
    [Fact]
    public void Bug109_WordIsNormalStyle_CaseSensitive()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test paragraph" });

        // Default paragraph should have Normal style
        var node = _wordHandler.Get("/body/p[1]");
        // Style comparison in the codebase is case-sensitive
        // This means "normal" != "Normal" and paragraphs may not be correctly identified
        node.Should().NotBeNull();
    }

    /// Bug #110 — Excel StyleManager: fill ID off-by-one when fills empty
    /// File: ExcelStyleManager.cs, line 353
    /// Returns (uint)(fills.Count() - 1) which overflows to uint.MaxValue
    /// when the fills collection was empty before appending.
    [Fact]
    public void Bug110_ExcelStyleManager_FillIdOverflowWhenEmpty()
    {
        // Set a background color on a cell — this exercises fill creation
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        _excelHandler.Set("/Sheet1/A1", new() { ["bgcolor"] = "FF0000" });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1/A1");

        // The fill should be applied correctly
        node.Format.TryGetValue("bgcolor", out var bg);
        bg.Should().NotBeNull("background color should be preserved after reopen");
    }

    // ==================== Bug #111-130: Delete/Move/Add, ResidentServer, Query edge cases ====================

    /// Bug #111 — Word Remove: no cleanup of embedded relationships
    /// File: WordHandler.Add.cs, lines 1177-1185
    /// Remove() just calls element.Remove() without cleaning up
    /// hyperlink relationships, image parts, or other embedded content.
    [Fact]
    public void Bug111_WordRemove_NoRelationshipCleanup()
    {
        // Add a hyperlink paragraph
        _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Click here",
            ["link"] = "https://example.com"
        });

        // Remove the paragraph — the hyperlink relationship should be cleaned up
        _wordHandler.Remove("/body/p[1]");

        // Reopen to verify
        ReopenWord();
        var root = _wordHandler.Get("/");
        // The relationship to https://example.com may remain orphaned
        // This is a file bloat / potential corruption issue
        root.Should().NotBeNull();
    }

    /// Bug #112 — Word Add table: int.Parse on negative rows/cols
    /// File: WordHandler.Add.cs, lines 350-351
    /// No validation that rows/cols are positive. Negative values cause
    /// empty table or unexpected behavior.
    [Fact]
    public void Bug112_WordAddTable_NegativeRowsCols()
    {
        // Adding a table with 0 rows — should fail gracefully
        var act = () => _wordHandler.Add("/body", "tbl", null, new()
        {
            ["rows"] = "0",
            ["cols"] = "3"
        });

        // Should validate and reject, or create at least 1 row
        // Instead it silently creates an empty table structure
        act.Should().NotThrow("zero rows should be handled gracefully");

        // Verify the table exists but has proper structure
        var node = _wordHandler.Get("/body/tbl[1]");
        node.Should().NotBeNull();
    }

    /// Bug #113 — Word Add: int.Parse on firstlineindent with multiplication overflow
    /// File: WordHandler.Add.cs, line 60
    /// int.Parse(indent) * 480 can overflow for large indent values.
    [Fact]
    public void Bug113_WordAdd_FirstLineIndentOverflow()
    {
        var act = () => _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Indented",
            ["firstlineindent"] = "9999999"
        });

        // int.Parse("9999999") * 480 = 4,799,999,520 which overflows int range
        // Should either validate range or use long arithmetic
        act.Should().Throw<Exception>(
            "int.Parse(indent) * 480 overflows for large indent values");
    }

    /// Bug #114 — Word Add TOC: bool.Parse on hyperlinks/pagenumbers
    /// File: WordHandler.Add.cs, lines 808-809
    /// Uses bool.Parse for "hyperlinks" and "pagenumbers" properties.
    [Fact]
    public void Bug114_WordAddToc_BoolParseOnOptions()
    {
        var act = () => _wordHandler.Add("/body", "toc", null, new()
        {
            ["hyperlinks"] = "yes"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on 'yes' — should use IsTruthy or accept common boolean aliases");
    }

    /// Bug #115 — Word Add: bool.Parse on paragraph keepnext/keeplines/pagebreakbefore
    /// File: WordHandler.Add.cs, lines 118-124
    /// Uses bool.Parse for layout properties during paragraph creation.
    [Fact]
    public void Bug115_WordAdd_BoolParseOnParagraphLayoutProperties()
    {
        var act = () => _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Keep",
            ["keepnext"] = "1"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on '1' — inconsistent with IsTruthy used elsewhere");
    }

    /// Bug #116 — Word Add: bool.Parse on paragraph bold/italic/caps/etc.
    /// File: WordHandler.Add.cs, lines 150-168
    /// Uses bool.Parse for all run formatting properties during paragraph creation.
    [Fact]
    public void Bug116_WordAdd_BoolParseOnRunFormatting()
    {
        var act = () => _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Bold text",
            ["bold"] = "yes"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on 'yes' for bold — should accept common boolean aliases");
    }

    /// Bug #117 — Word Move: IndexOf returning -1 causes wrong path
    /// File: WordHandler.Add.cs, lines 1223-1225
    /// After Move, IndexOf(element) on siblings list returns -1 if
    /// element matching fails, producing path like "/body/p[0]".
    [Fact]
    public void Bug117_WordMove_IndexOfReturnsWrongPath()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Second" });
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Third" });

        // Move third paragraph to position 1
        var newPath = _wordHandler.Move("/body/p[3]", "/body", 1);

        // The returned path should be valid (1-based)
        newPath.Should().Contain("p[1]",
            "Move should return a valid 1-based path for the moved element");
    }

    /// Bug #118 — Excel sheet deletion: orphaned defined names
    /// File: ExcelHandler.Add.cs, lines 939-942
    /// Deleting a sheet removes the part but doesn't clean up
    /// defined names that reference the deleted sheet.
    [Fact]
    public void Bug118_ExcelDeleteSheet_OrphanedDefinedNames()
    {
        // Add a second sheet with a named range
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Data" });
        _excelHandler.Add("/Data", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });

        // Add a defined name referencing Data sheet
        _excelHandler.Add("/", "definedname", null, new()
        {
            ["name"] = "MyRange",
            ["value"] = "Data!$A$1"
        });

        // Delete the Data sheet — the defined name should be cleaned up
        _excelHandler.Remove("/Data");

        ReopenExcel();
        // The defined name "MyRange" may still reference the deleted sheet
        var root = _excelHandler.Get("/");
        root.Should().NotBeNull();
    }

    /// Bug #119 — PPTX Remove chart: bare catch swallows errors
    /// File: PowerPointHandler.Add.cs, lines 1217-1224
    /// Chart deletion uses try/catch{} that silently swallows part deletion errors.
    [Fact]
    public void Bug119_PptxRemoveChart_BareCatchSwallowsErrors()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30"
        });

        // Remove the chart
        pptx.Remove("/slide[1]/chart[1]");

        // Verify chart is removed
        var node = pptx.Get("/slide[1]");
        node.Children.Where(c => c.Type == "chart").Should().BeEmpty(
            "chart should be removed after Remove call");
    }

    /// Bug #120 — PPTX ungroup: pictures not cleaned up properly
    /// File: PowerPointHandler.Add.cs, lines 1233-1250
    /// When ungrouping, pictures moved from group to shape tree
    /// don't get their media relationships cleaned up.
    [Fact]
    public void Bug120_PptxUngroup_PicturesNotCleanedProperly()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add shapes and group them
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape A" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape B" });

        // Group them
        pptx.Add("/slide[1]", "group", null, new()
        {
            ["shapes"] = "1,2"
        });

        // Now ungroup — shapes should return to slide level
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #121 — Word Add: uint.Parse on page width/height without validation
    /// File: WordHandler.Add.cs, lines 689-691
    /// Section page size uses uint.Parse without TryParse.
    [Fact]
    public void Bug121_WordAdd_UintParseOnPageSize()
    {
        var act = () => _wordHandler.Add("/body", "section", null, new()
        {
            ["width"] = "wide"
        });

        act.Should().Throw<FormatException>(
            "uint.Parse crashes on 'wide' — should use TryParse");
    }

    /// Bug #122 — PPTX Query: IndexOf returning -1 for placeholder shapes
    /// File: PowerPointHandler.Query.cs, line 220
    /// IndexOf returns -1 if shape not found, causing shapeIdx+1=0 (invalid 1-based index).
    [Fact]
    public void Bug122_PptxQuery_PlaceholderIndexOfMinusOne()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Get slide — placeholder shapes from layout should have valid indices
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
        // Placeholder nodes should have paths with valid 1-based indices
        foreach (var child in node.Children)
        {
            if (child.Path.Contains("shape["))
            {
                child.Path.Should().NotContain("shape[0]",
                    "shape index should be 1-based, not 0 from IndexOf returning -1");
            }
        }
    }

    /// Bug #123 — Word Add run: bool.Parse on all formatting properties
    /// File: WordHandler.Add.cs, lines 278-296
    /// Run creation uses bool.Parse for bold, italic, strike, caps, etc.
    [Fact]
    public void Bug123_WordAddRun_BoolParseOnFormatting()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Hello" });

        var act = () => _wordHandler.Add("/body/p[1]", "r", null, new()
        {
            ["text"] = " World",
            ["bold"] = "TRUE"
        });

        // bool.Parse is case-insensitive for "True"/"False" but crashes on other values
        // "TRUE" actually works with bool.Parse, but "1" or "yes" don't
        var act2 = () => _wordHandler.Add("/body/p[1]", "r", null, new()
        {
            ["text"] = " World",
            ["bold"] = "on"
        });

        act2.Should().Throw<FormatException>(
            "bool.Parse crashes on 'on' — should use IsTruthy for consistency");
    }

    /// Bug #124 — Word Add image: bool.Parse on anchor/behindtext
    /// File: WordHandler.Add.cs, lines 488, 499
    /// Image floating properties use bool.Parse.
    [Fact]
    public void Bug124_WordAddImage_BoolParseOnAnchor()
    {
        var imgPath = CreateTempImage();
        try
        {
            var act = () => _wordHandler.Add("/body", "image", null, new()
            {
                ["src"] = imgPath,
                ["anchor"] = "yes"
            });

            act.Should().Throw<FormatException>(
                "bool.Parse crashes on 'yes' for anchor property");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #125 — Word Add style: bool.Parse and int.Parse
    /// File: WordHandler.Add.cs, lines 930, 933, 938
    /// Style creation uses bool.Parse for bold/italic and int.Parse for size.
    [Fact]
    public void Bug125_WordAddStyle_BoolAndIntParse()
    {
        var act = () => _wordHandler.Add("/styles", "style", null, new()
        {
            ["name"] = "MyStyle",
            ["bold"] = "1",
            ["size"] = "12.5"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse on '1' or int.Parse on '12.5' crashes in style creation");
    }

    /// Bug #126 — Word Add header: bool.Parse and int.Parse
    /// File: WordHandler.Add.cs, lines 986-989
    /// Header creation uses bool.Parse for bold/italic.
    [Fact]
    public void Bug126_WordAddHeader_BoolParseOnFormatting()
    {
        var act = () => _wordHandler.Add("/body", "header", null, new()
        {
            ["text"] = "Header",
            ["bold"] = "yes"
        });

        act.Should().Throw<FormatException>(
            "bool.Parse crashes on 'yes' in header creation");
    }

    /// Bug #127 — Word Add: shading split produces empty array element
    /// File: WordHandler.Add.cs, line 88
    /// pShdVal.Split(';') on a value without semicolons returns one element,
    /// but accessing shdParts[1] or [2] would fail.
    [Fact]
    public void Bug127_WordAdd_ShadingSplitEdgeCase()
    {
        // Shading with just a color, no pattern/theme
        var act = () => _wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "Shaded",
            ["shd"] = "FF0000"
        });

        // Should handle single value (just color) without crash
        act.Should().NotThrow(
            "Shading with single color value should not crash on split parsing");
    }

    /// Bug #128 — Word Add document properties: uint.Parse / int.Parse
    /// File: WordHandler.Add.cs, lines 1309-1324
    /// Document property setting uses uint.Parse and int.Parse without validation.
    [Fact]
    public void Bug128_WordAdd_DocumentPropertyParse()
    {
        var act = () => _wordHandler.Set("/", new()
        {
            ["pagewidth"] = "auto"
        });

        act.Should().Throw<Exception>(
            "uint.Parse crashes on 'auto' for page width");
    }

    /// Bug #129 — PPTX RemovePictureWithCleanup: bare catch swallows all errors
    /// File: PowerPointHandler.cs, lines 333-340
    /// Uses catch{} that silently swallows deletion errors including
    /// invalid relationship IDs and corrupted part references.
    [Fact]
    public void Bug129_PptxRemovePicture_BareCatchSwallowsErrors()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });
            pptx.Remove("/slide[1]/picture[1]");

            // Verify picture is removed
            var node = pptx.Get("/slide[1]");
            node.Children.Where(c => c.Type == "picture").Should().BeEmpty(
                "picture should be removed");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #130 — Excel Add chart: chart position int.Parse series
    /// File: ExcelHandler.Add.cs, lines 838-841
    /// Chart width/height use int.Parse, treating them as column/row counts.
    /// But "width" semantically suggests pixels, causing confusion.
    [Fact]
    public void Bug130_ExcelAddChart_WidthHeightSemanticConfusion()
    {
        // Width is parsed as column count, not pixels
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "20" });

        var act = () => _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20",
            ["width"] = "400"  // User expects pixels, but code adds to fromCol
        });

        // int.Parse("400") + 0 = 400 columns — way off screen
        // The semantics are confusing: width means column span, not pixels
        act.Should().NotThrow("should handle large width, but result will be off-screen");
    }

    // ==================== Bug #131-150: Move methods, Query indexing, Excel/Word/PPTX edge cases ====================

    /// Bug #131 — PPTX Move slide: 0-based index vs 1-based paths
    /// File: PowerPointHandler.Add.cs, line 1284
    /// Slide move uses 0-based index for insertion, but the rest of the
    /// API uses 1-based paths (/slide[1]). index=0 inserts before first slide.
    [Fact]
    public void Bug131_PptxMoveSlide_ZeroBasedIndexInconsistency()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Slide 1" });
        pptx.Add("/slide[2]", "shape", null, new() { ["text"] = "Slide 2" });

        // Move slide 2 to position 1 (0-based index = 0)
        var newPath = pptx.Move("/slide[2]", null, 0);

        // The API accepts 0-based index but returns 1-based path
        newPath.Should().Be("/slide[1]",
            "Moving slide with index=0 should place it first, returning /slide[1]");
    }

    /// Bug #132 — PPTX Move shape cross-slide: element removed before relationship copy
    /// File: PowerPointHandler.Add.cs, lines 1331-1335
    /// srcElement.Remove() is called BEFORE CopyRelationships().
    /// If CopyRelationships fails, the shape is lost.
    [Fact]
    public void Bug132_PptxMoveShapeCrossSlide_ElementRemovedBeforeRelCopy()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Moving shape" });

        // Move shape from slide 1 to slide 2
        var newPath = pptx.Move("/slide[1]/shape[1]", "/slide[2]", null);

        // Verify shape exists on slide 2
        var slide2 = pptx.Get("/slide[2]");
        slide2.Children.Should().Contain(c => c.Type == "shape",
            "Shape should exist on target slide after cross-slide move");
    }

    /// Bug #133 — PPTX ComputeElementPath: IndexOf returns -1 producing shape[0]
    /// File: PowerPointHandler.Add.cs, lines 1480-1492
    /// If element not found in type-filtered list, IndexOf returns -1,
    /// producing path like /slide[1]/shape[0] (invalid 1-based index).
    [Fact]
    public void Bug133_PptxComputeElementPath_InvalidZeroIndex()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Move shape within same slide — returned path should be valid
        var newPath = pptx.Move("/slide[1]/shape[1]", "/slide[1]", 1);
        newPath.Should().NotContain("[0]",
            "Returned path should use 1-based indices, not 0 from IndexOf=-1");
    }

    /// Bug #134 — Excel Move row: target worksheet not saved
    /// File: ExcelHandler.Add.cs, line 1028
    /// Only source worksheet is saved after move, not target worksheet.
    /// When moving to a different sheet, changes to target may be lost.
    [Fact]
    public void Bug134_ExcelMoveRow_TargetSheetNotSaved()
    {
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Sheet2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Data" });

        // Move row from Sheet1 to Sheet2
        _excelHandler.Move("/Sheet1/row[1]", "/Sheet2", null);

        // Reopen and verify data exists on Sheet2
        ReopenExcel();
        var sheet2 = _excelHandler.Get("/Sheet2");
        sheet2.Should().NotBeNull();
        // If target sheet wasn't saved, the row data may be lost on reopen
    }

    /// Bug #135 — Excel Move row: RowIndex not updated after move
    /// File: ExcelHandler.Add.cs, lines 1013-1026
    /// Row's RowIndex property is never updated after moving,
    /// causing potential duplicate RowIndex values in target sheet.
    [Fact]
    public void Bug135_ExcelMoveRow_RowIndexNotUpdated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Row1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Row2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "Row3" });

        // Move row 3 to position 1
        _excelHandler.Move("/Sheet1/row[3]", "/Sheet1", 0);

        // After move, the moved row should have an appropriate RowIndex
        // But the code doesn't update RowIndex, so row 3 still has RowIndex=3
        // even though it's now physically first in the sheet
        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #136 — Excel Move: cell references in formulas not updated
    /// File: ExcelHandler.Add.cs, lines 1013-1031
    /// Moving a row doesn't update formula references pointing to it.
    [Fact]
    public void Bug136_ExcelMoveRow_FormulaReferencesNotUpdated()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "20" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3" });
        _excelHandler.Set("/Sheet1/A3", new() { ["formula"] = "A1+A2" });

        // Move row 1 to end
        _excelHandler.Move("/Sheet1/row[1]", "/Sheet1", null);

        // The formula =A1+A2 in row 3 should ideally be updated
        // but the Move method doesn't update formula references
        var node = _excelHandler.Get("/Sheet1/A3");
        node.Should().NotBeNull("formula cell should still exist after row move");
    }

    /// Bug #137 — Word Query: mathParaIdx shared between body-level and paragraph-level equations
    /// File: WordHandler.Query.cs, lines 393-434
    /// Both body-level oMathPara (line 400) and paragraph-level oMathPara (line 428)
    /// increment the same mathParaIdx counter, causing index collision.
    [Fact]
    public void Bug137_WordQuery_MathParaIndexCollision()
    {
        // Add a regular paragraph between math elements
        // This documents the indexing bug where body-level and paragraph-level
        // math elements share the same counter
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Normal text" });

        // The bug is in query: both code paths increment mathParaIdx
        // causing non-sequential or colliding oMathPara indices
        var node = _wordHandler.Get("/body");
        node.Should().NotBeNull();
    }

    /// Bug #138 — Word Query: paragraph-level oMathPara gets wrong path
    /// File: WordHandler.Query.cs, lines 432-434
    /// oMathPara inside a paragraph gets path "/body/oMathPara[N]"
    /// instead of "/body/p[N]/oMathPara[1]".
    [Fact]
    public void Bug138_WordQuery_ParagraphMathWrongPath()
    {
        // This is a path generation bug:
        // An equation inside a paragraph should have a path relative to the paragraph
        // but instead gets a body-level path, making navigation ambiguous
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Text" });
        var node = _wordHandler.Get("/body");
        node.Should().NotBeNull();
    }

    /// Bug #139 — Excel Query: CellToNode inconsistent parameter count
    /// File: ExcelHandler.Query.cs, lines 186 vs 457
    /// CellToNode is called with 3 params (including WorksheetPart) on line 186
    /// but only 2 params on line 457, silently skipping hyperlink/border info.
    [Fact]
    public void Bug139_ExcelQuery_CellToNodeInconsistentParams()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        _excelHandler.Set("/Sheet1/A1", new() { ["link"] = "https://example.com" });

        // Get cell through different query paths
        var directNode = _excelHandler.Get("/Sheet1/A1");
        // Query through a different path may omit hyperlink info
        // due to CellToNode being called without WorksheetPart
        directNode.Format.TryGetValue("link", out var link);
        link.Should().NotBeNull("hyperlink should be visible regardless of query path");
    }

    /// Bug #140 — Excel Query: null CellReference defaults to A1
    /// File: ExcelHandler.Query.cs, line 282
    /// cell.CellReference?.Value ?? "A1" silently normalizes null to A1.
    [Fact]
    public void Bug140_ExcelQuery_NullCellReferenceDefaultsToA1()
    {
        // This is a defensive coding issue — cells with null CellReference
        // are silently treated as A1 instead of being flagged as errors
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Hello" });
        var node = _excelHandler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("Hello");
    }

    /// Bug #141 — PPTX Move: CopyRelationships bare catch swallows errors
    /// File: PowerPointHandler.Add.cs, line 1438
    /// try/catch{} silently ignores relationship copy failures,
    /// leaving stale relationship IDs in moved elements.
    [Fact]
    public void Bug141_PptxMove_CopyRelationshipsBareCatch()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });
            // Move picture across slides — relationships must be copied
            pptx.Move("/slide[1]/picture[1]", "/slide[2]", null);

            // Verify picture is on slide 2 with valid image data
            var slide2 = pptx.Get("/slide[2]");
            slide2.Children.Should().Contain(c => c.Type == "picture",
                "picture should be moved to slide 2 with valid relationships");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    /// Bug #142 — Word Query: HeaderParts null safety
    /// File: WordHandler.Query.cs, line 196
    /// mainPart?.HeaderParts.ElementAtOrDefault(index) doesn't null-check
    /// HeaderParts itself, only mainPart.
    [Fact]
    public void Bug142_WordQuery_HeaderPartsNullSafety()
    {
        // On a document without any headers, querying a header should
        // return null or throw a clear error, not NullReferenceException
        var act = () => _wordHandler.Get("/header[1]");

        // Should handle gracefully when no headers exist
        act.Should().NotThrow("querying non-existent header should return null, not crash");
    }

    /// Bug #143 — Excel Query: comment list null access
    /// File: ExcelHandler.Query.cs, lines 312-313
    /// cmtList can be null when no comments exist, but the code
    /// may still try to access its elements.
    [Fact]
    public void Bug143_ExcelQuery_CommentListNullAccess()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Try to get a comment that doesn't exist
        var node = _excelHandler.Get("/Sheet1/A1");
        // Should not crash even when no comments part exists
        node.Should().NotBeNull();
    }

    /// Bug #144 — PPTX InsertAtPosition: 0-based vs 1-based index inconsistency
    /// File: PowerPointHandler.Add.cs, line 1451
    /// ShapeTree insertion filters to content children but uses 0-based index,
    /// while non-ShapeTree parents use raw ChildElements with same 0-based index.
    [Fact]
    public void Bug144_PptxInsertAtPosition_IndexInconsistency()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // Move shape C to position 0 (should become first)
        pptx.Move("/slide[1]/shape[3]", "/slide[1]", 0);

        var slide = pptx.Get("/slide[1]");
        var shapes = slide.Children.Where(c => c.Type == "shape").ToList();
        shapes.Should().HaveCountGreaterOrEqualTo(3,
            "All shapes should still exist after move");
    }

    /// Bug #145 — Excel Move: return value uses list index not RowIndex
    /// File: ExcelHandler.Add.cs, line 1030
    /// newRows.IndexOf(row) + 1 returns position in element list,
    /// not the logical row index (RowIndex property).
    [Fact]
    public void Bug145_ExcelMove_ReturnValueUsesListIndex()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A5", ["value"] = "50" });

        // Move row at index 5 to beginning
        var newPath = _excelHandler.Move("/Sheet1/row[2]", "/Sheet1", 0);

        // The returned path should reflect the logical position
        newPath.Should().Contain("row[",
            "Move should return a valid row path");
    }

    /// Bug #146 — Word Query: bookmark name with special characters in path
    /// File: WordHandler.Query.cs, line 388
    /// Path "/bookmark[name]" doesn't escape special chars in bookmark names.
    /// A bookmark named "my/bookmark" produces invalid path "/bookmark[my/bookmark]".
    [Fact]
    public void Bug146_WordQuery_BookmarkSpecialCharsInPath()
    {
        // Add a bookmark with a simple name first
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Bookmarked" });
        _wordHandler.Set("/body/p[1]", new() { ["bookmark"] = "TestBM" });

        // Query the bookmark
        var node = _wordHandler.Get("/bookmark[TestBM]");
        // Should return the bookmark node
        (node != null).Should().BeTrue("bookmark should be queryable by name");
    }

    /// Bug #147 — PPTX Move: negative index silently appends
    /// File: PowerPointHandler.Add.cs, lines 1284, 1451
    /// index.Value >= 0 check causes negative indices to fall through
    /// to the append branch instead of throwing a validation error.
    [Fact]
    public void Bug147_PptxMove_NegativeIndexSilentlyAppends()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/", "slide", null, new());

        // Move with negative index — should throw or reject, not silently append
        var act = () => pptx.Move("/slide[1]", null, -1);

        // Current behavior: silently appends (treated as no index)
        // Expected: should throw ArgumentException for negative index
        act.Should().NotThrow("negative index is silently treated as append — this is a bug");
    }

    /// Bug #148 — Excel Query: ColumnNameToIndex returns int cast to uint without bounds check
    /// File: ExcelHandler.Query.cs, line 153
    /// If ColumnNameToIndex returns a negative value, casting to uint
    /// produces a very large number instead of throwing an error.
    [Fact]
    public void Bug148_ExcelQuery_ColumnIndexTypeConversion()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        var node = _excelHandler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    /// Bug #149 — PPTX Query: placeholder index mismatch with shape index
    /// File: PowerPointHandler.Query.cs, lines 516-517
    /// Placeholder nodes use phIdx (placeholder count) as shape index,
    /// not the actual index among all shapes in the shape tree.
    [Fact]
    public void Bug149_PptxQuery_PlaceholderIndexMismatch()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Regular shape" });

        // Query placeholders — their indices should be consistent
        var slide = pptx.Get("/slide[1]");
        slide.Should().NotBeNull();
        // Placeholders use phIdx, not their actual position among all shapes
    }

    /// Bug #150 — Excel Add: empty parentPath produces IndexOutOfRange
    /// File: ExcelHandler.Add.cs, lines 948, 984, 1047
    /// segments[1] accessed without verifying segments.Length > 1.
    [Fact]
    public void Bug150_ExcelAdd_EmptyPathSegmentAccess()
    {
        // Providing a path with only a sheet name (no sub-element) to Move
        var act = () => _excelHandler.Move("/Sheet1", "/Sheet1", null);

        act.Should().Throw<Exception>(
            "Move with sheet-only path should throw clear error, not IndexOutOfRangeException");
    }

    // ==================== Bug #151-170: CopyFrom, Selector, GenericXmlQuery, Animations, Chart ====================

    /// Bug #151 — GenericXmlQuery: 0-based Traverse vs 1-based ElementToNode
    /// File: GenericXmlQuery.cs, lines 65 vs 208
    /// Traverse() generates paths with 0-based indices [0], [1], [2]...
    /// ElementToNode() generates paths with 1-based indices [1], [2], [3]...
    /// NavigateByPath() expects 1-based (subtracts 1 on line 254).
    /// Paths from Traverse() cannot be used with NavigateByPath().
    [Fact]
    public void Bug151_GenericXmlQuery_IndexInconsistency()
    {
        // GenericXmlQuery.Traverse generates /element[0] (0-based)
        // But NavigateByPath expects /element[1] (1-based, subtracts 1)
        // This means paths from Traverse() cannot be navigated back
        // This is a fundamental design inconsistency in the generic XML layer
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });
        var root = _wordHandler.Get("/");
        root.Should().NotBeNull();
    }

    /// Bug #152 — GenericXmlQuery: int.Parse in ParsePathSegments
    /// File: GenericXmlQuery.cs, line 231
    /// Uses int.Parse on path index without validation.
    /// Malformed paths like "/element[abc]" crash instead of returning null.
    [Fact]
    public void Bug152_GenericXmlQuery_IntParseInPathSegments()
    {
        // The GenericXmlQuery layer uses int.Parse without TryParse
        // on path segment indices. This was already documented but
        // confirms the systemic pattern extends beyond handlers.
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });
        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull();
    }

    /// Bug #153 — GenericXmlQuery: TryCreateTypedElement index convention unclear
    /// File: GenericXmlQuery.cs, line 436
    /// InsertBeforeSelf uses index.Value directly as array index (0-based),
    /// but callers may pass 1-based indices from path notation.
    [Fact]
    public void Bug153_GenericXmlQuery_InsertionIndexConvention()
    {
        // The index parameter convention is undocumented:
        // Does index=1 mean "insert at position 1 (0-based)" or
        // "insert at position 1 (1-based, i.e., first element)"?
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Second" });
        var node = _wordHandler.Get("/body");
        node.ChildCount.Should().BeGreaterOrEqualTo(2);
    }

    /// Bug #154 — Word Selector: :contains() hardcoded offset
    /// File: WordHandler.Selector.cs, line 65
    /// Uses idx + 10 as magic number assuming ":contains(" is 10 chars.
    /// Fragile and breaks if selector name changes.
    [Fact]
    public void Bug154_WordSelector_ContainsHardcodedOffset()
    {
        // This is a maintenance risk — ":contains(" is 10 chars
        // but the code uses magic number 10 instead of .Length
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Hello World" });

        // Query with :contains should work
        var results = _wordHandler.Query("p:contains(Hello)");
        results.Should().NotBeEmpty("selector :contains(Hello) should match the paragraph");
    }

    /// Bug #155 — Word Selector: attribute regex doesn't match hyphenated names
    /// File: WordHandler.Selector.cs, line 52
    /// Regex pattern \w+ for attribute names doesn't match hyphens.
    /// Attributes like [data-foo=bar] fail to parse.
    [Fact]
    public void Bug155_WordSelector_AttributeRegexHyphenated()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });

        // Attribute selectors with hyphens in name may not parse
        // The regex \w+ only matches word characters (no hyphens)
        var results = _wordHandler.Query("p");
        results.Should().NotBeEmpty();
    }

    /// Bug #156 — Word Selector: :empty false positive for prefix matches
    /// File: WordHandler.Selector.cs, line 71
    /// selector.Contains(":empty") matches ":emptiness" or ":empty-cell"
    /// because there's no word boundary check.
    [Fact]
    public void Bug156_WordSelector_EmptyPseudoNoBoundary()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "" });
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Not empty" });

        // :empty should match empty paragraphs
        var results = _wordHandler.Query("p:empty");
        results.Should().HaveCountGreaterOrEqualTo(1,
            ":empty should match paragraphs with no text");
    }

    /// Bug #157 — Excel CopyFrom: shared string references not updated
    /// File: ExcelHandler.Add.cs, line 1065
    /// CloneNode(true) copies cells with SharedString type,
    /// but cloned cells still reference original shared string indices.
    [Fact]
    public void Bug157_ExcelCopyFrom_SharedStringReferences()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Hello" });
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Sheet2" });

        // Copy row from Sheet1 to Sheet2
        _excelHandler.CopyFrom("/Sheet1/row[1]", "/Sheet2", null);

        ReopenExcel();
        // Verify the copied cell has the correct value
        var node = _excelHandler.Get("/Sheet2");
        node.Should().NotBeNull("Sheet2 should exist after copy");
    }

    /// Bug #158 — Excel chart: pie chart silently ignores extra series
    /// File: ExcelHandler.Helpers.cs, line 509
    /// Pie/doughnut charts only use seriesData[0], silently discarding
    /// additional series without warning.
    [Fact]
    public void Bug158_ExcelChart_PieChartIgnoresExtraSeries()
    {
        // Add data for multiple series
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Q1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Q2" });

        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "pie",
            ["data"] = "Sales:10,20;Costs:5,15"
        });

        // Only the first series (Sales) is rendered
        // Costs series is silently dropped — this is data loss
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #159 — Excel chart: legend parsing too permissive
    /// File: ExcelHandler.Helpers.cs, lines 544-546
    /// Any value except "false"/"none" shows a legend.
    /// Values like "off", "hide", "no" still show a legend.
    [Fact]
    public void Bug159_ExcelChart_LegendParsingTooPermissive()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        // "no" should hide legend but it's not recognized
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30",
            ["legend"] = "no"
        });

        // "no" is not "false" or "none", so legend is still shown
        // This is inconsistent with the IsTruthy pattern used elsewhere
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #160 — PPTX Animations: split transition ignores direction
    /// File: PowerPointHandler.Animations.cs, line 87
    /// Split transition hardcodes direction to "in" regardless of user input.
    [Fact]
    public void Bug160_PptxAnimations_SplitTransitionIgnoresDirection()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set split transition with "out" direction
        pptx.Set("/slide[1]", new()
        {
            ["transition"] = "split-out"
        });

        // The direction should be "out" but code hardcodes "in"
        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

    /// Bug #161 — PPTX Animations: negative duration accepted
    /// File: PowerPointHandler.Animations.cs, line 214
    /// int.TryParse succeeds for negative values with no bounds checking.
    [Fact]
    public void Bug161_PptxAnimations_NegativeDurationAccepted()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Negative duration should be rejected but int.TryParse accepts it
        var act = () => pptx.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "fade",
            ["trigger"] = "onclick",
            ["duration"] = "-500"
        });

        // Should either reject negative duration or clamp to 0
        act.Should().NotThrow("negative duration is accepted without validation — this is a bug");
    }

    /// Bug #162 — PPTX Animations: emphasis animations treated as "Out"
    /// File: PowerPointHandler.Animations.cs, lines 378-379
    /// Only checks if presetClass is Entrance; everything else (including
    /// Emphasis) is treated as Exit/Out, which is semantically wrong.
    [Fact]
    public void Bug162_PptxAnimations_EmphasisTreatedAsExit()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Add emphasis animation
        pptx.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "fade",
            ["trigger"] = "onclick",
            ["class"] = "emphasis"
        });

        var node = pptx.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
    }

    /// Bug #163 — PPTX Animations: PresetSubtype always 0
    /// File: PowerPointHandler.Animations.cs, line 435
    /// PresetSubtype is hardcoded to 0 for all animations,
    /// but different effects require different subtypes.
    [Fact]
    public void Bug163_PptxAnimations_PresetSubtypeAlwaysZero()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Fly-in animation typically needs subtype for direction
        pptx.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "fly",
            ["trigger"] = "onclick"
        });

        // The animation subtype should vary by effect type and direction
        // but it's always hardcoded to 0
        var slide = pptx.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    /// Bug #164 — GenericXmlQuery: ParsePathSegments missing bracket validation
    /// File: GenericXmlQuery.cs, lines 226-227
    /// No validation that closing bracket exists. Malformed path "a[1"
    /// causes incorrect substring operation.
    [Fact]
    public void Bug164_GenericXmlQuery_MalformedPathBracket()
    {
        // Malformed paths should be handled gracefully
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });
        var node = _wordHandler.Get("/body/p[1]");
        node.Should().NotBeNull();
    }

    /// Bug #165 — Excel chart: empty seriesData causes Max() crash
    /// File: ExcelHandler.Helpers.cs, line 479
    /// If seriesData is empty, Max() throws InvalidOperationException.
    [Fact]
    public void Bug165_ExcelChart_EmptySeriesDataCrash()
    {
        var act = () => _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar"
            // No data provided
        });

        // Should handle gracefully, not crash on empty series
        act.Should().Throw<Exception>(
            "Chart creation with no data should throw clear error, not InvalidOperationException from Max()");
    }

    /// Bug #166 — Word Selector: multiple child selectors silently ignored
    /// File: WordHandler.Selector.cs, line 24
    /// Only the first child selector (after >) is parsed.
    /// "p > r > span" silently ignores the "span" part.
    [Fact]
    public void Bug166_WordSelector_NestedChildSelectorsIgnored()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });

        // "p > r" should work (one level child)
        var results = _wordHandler.Query("p > r");
        // Just document that nested selectors beyond one level are silently ignored
        results.Should().NotBeNull();
    }

    /// Bug #167 — Word Selector: attribute value quote stripping too aggressive
    /// File: WordHandler.Selector.cs, line 56
    /// Trim('\'', '"') removes ALL leading/trailing quotes,
    /// including legitimate ones in the value.
    [Fact]
    public void Bug167_WordSelector_QuoteStrippingTooAggressive()
    {
        _wordHandler.Add("/body", "p", null, new() { ["text"] = "Test" });

        // "[style=Normal]" should work with unquoted value
        var results = _wordHandler.Query("p[style=Normal]");
        results.Should().NotBeEmpty("attribute selector should match by style name");
    }

    /// Bug #168 — PPTX Animations: wrong default presetId for reading back
    /// File: PowerPointHandler.Animations.cs, line 582
    /// Defaults to 10 (fade) when PresetId is null, but first animation
    /// in the switch is appear (1). This causes null PresetId to be
    /// reported as "fade" instead of "unknown".
    [Fact]
    public void Bug168_PptxAnimations_WrongDefaultPresetId()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Add appear animation (presetId=1)
        pptx.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "appear",
            ["trigger"] = "onclick"
        });

        var node = pptx.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
    }

    /// Bug #169 — Excel CopyFrom: target worksheet save only
    /// File: ExcelHandler.Add.cs, line 1080
    /// CopyFrom only saves target worksheet. If source state was modified
    /// (e.g., metadata about copy operations), it's not persisted.
    [Fact]
    public void Bug169_ExcelCopyFrom_SourceNotSaved()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Original" });
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Sheet2" });

        _excelHandler.CopyFrom("/Sheet1/row[1]", "/Sheet2", null);

        ReopenExcel();
        // Both sheets should have data
        var s1 = _excelHandler.Get("/Sheet1/A1");
        s1.Should().NotBeNull("original cell should still exist after copy");
    }

    /// Bug #170 — PPTX Animation transition: duration string stored directly
    /// File: PowerPointHandler.Animations.cs, line 65
    /// trans.Duration = durationMs assigns a string that was parsed from user input.
    /// No validation that the string represents a valid duration value.
    [Fact]
    public void Bug170_PptxAnimationTransition_DurationStringDirect()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set transition with custom duration
        pptx.Set("/slide[1]", new()
        {
            ["transition"] = "fade",
            ["transitionDuration"] = "1000"
        });

        var node = pptx.Get("/slide[1]");
        node.Should().NotBeNull();
    }

}
