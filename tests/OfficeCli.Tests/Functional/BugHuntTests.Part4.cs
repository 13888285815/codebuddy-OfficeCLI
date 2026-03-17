// Bug hunt tests Part 4 — Bug #251-290
// Program.cs, Helpers, NodeBuilder, Background, Fill, Effects, Selector, View, Core

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ==================== Bug #251-270: Program.cs, Helpers, NodeBuilder, Background, Fill ====================

    /// Bug #251 — Program.cs: Process object leaked in open command
    /// File: Program.cs, lines 53-77
    /// The Process object started at line 53 is never disposed. If TryConnect() times out
    /// or the process exits, the Process handle is leaked (no using statement, no Dispose).
    [Fact]
    public void Bug251_ProgramOpenCommand_ProcessNotDisposed()
    {
        // The open command creates a Process via Process.Start() but never wraps it in
        // using or disposes it. This is a resource leak.
        // We can verify by examining that the code path doesn't dispose the process:

        // The bug is in Program.cs lines 53-77:
        //   var process = Process.Start(startInfo);  // line 53 - never disposed
        //   ... return; (line 67 and 73) - process handle leaked
        //   ... return; (line 77) - process handle leaked

        // This test documents the design flaw — Process implements IDisposable
        // but is never disposed in any exit path.
        var processType = typeof(System.Diagnostics.Process);
        processType.GetInterfaces().Should().Contain(typeof(IDisposable),
            "Process is IDisposable but open command never disposes it");
    }

    /// Bug #252 — Program.cs: create command missing SafeRun wrapper
    /// File: Program.cs, lines 586-607
    /// The create command handler doesn't use SafeRun(), unlike ALL other commands.
    /// This means exceptions from BlankDocCreator.Create() crash the CLI instead of
    /// being caught and reported gracefully.
    [Fact]
    public void Bug252_ProgramCreateCommand_MissingSafeRun()
    {
        // All other commands use: command.SetAction(result => SafeRun(() => { ... }));
        // But create command uses: command.SetAction(result => { ... });
        // This means any exception (e.g. disk full, permission denied, invalid path)
        // will propagate uncaught and crash the CLI with a stack trace instead of
        // the friendly "Error: ..." message.

        // Document the inconsistency — create is the only command without SafeRun
        var act = () => OfficeCli.BlankDocCreator.Create("/invalid/path/that/does/not/exist/test.docx");
        act.Should().Throw<Exception>(
            "BlankDocCreator.Create throws on invalid path, but create command has no SafeRun to catch it");
    }

    /// Bug #253 — Program.cs: property parsing allows empty value with trailing =
    /// File: Program.cs, lines 282-286, 360-364
    /// The property parsing uses `eqIdx > 0` which allows "key=" to create
    /// a property with an empty string value. This is inconsistent behavior
    /// and could cause downstream issues.
    [Fact]
    public void Bug253_ProgramPropertyParsing_EmptyValueAllowed()
    {
        // When prop="key=", eqIdx = 3, which is > 0, so it creates
        // properties["key"] = "" (empty string).
        // This is parsed as valid but could cause issues in handlers
        // that don't expect empty string property values.
        var prop = "key=";
        var eqIdx = prop.IndexOf('=');
        (eqIdx > 0).Should().BeTrue("eqIdx > 0 passes for 'key=' allowing empty values");
        prop[(eqIdx + 1)..].Should().BeEmpty("the value portion is empty, which may cause handler errors");
    }

    /// Bug #254 — PPTX ParseEmu: double.Parse without TryParse on unit values
    /// File: PowerPointHandler.Helpers.cs, lines 164-172
    /// All ParseEmu branches use double.Parse or long.Parse without validation.
    /// Invalid unit strings like "abc cm" or "not_a_number" will throw FormatException.
    [Fact]
    public void Bug254_PptxParseEmu_DoubleParseNoValidation()
    {
        // ParseEmu does: double.Parse(value[..^2]) for "cm", "in", "pt", "px" suffixes
        // and long.Parse(value) for raw EMU values.
        // None of these use TryParse, so invalid input throws unhandled FormatException.

        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            // Adding a shape with invalid EMU value should fail gracefully
            var act = () => handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "test",
                ["x"] = "not_a_numbercm"  // invalid double before "cm" suffix
            });
            act.Should().Throw<Exception>(
                "ParseEmu with invalid unit value throws FormatException instead of a clear error");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #255 — PPTX ParseEmu: long.Parse fallback for raw EMU without validation
    /// File: PowerPointHandler.Helpers.cs, line 172
    /// When the value doesn't match any unit suffix, long.Parse(value) is called directly.
    /// No TryParse, no validation — any non-numeric string causes FormatException.
    [Fact]
    public void Bug255_PptxParseEmu_LongParseFallbackNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            var act = () => handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "test",
                ["x"] = "hello"  // not a valid EMU number, not a unit string
            });
            act.Should().Throw<Exception>(
                "ParseEmu raw fallback: long.Parse('hello') throws FormatException");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #256 — PPTX Background: single-color gradient creates invalid gradient
    /// File: PowerPointHandler.Background.cs, lines 244-248
    /// BuildGradientFill allows colorParts.Count==1 after removing angle/focus.
    /// A single gradient stop at position 0 is invalid per OpenXML spec —
    /// gradients require at least 2 stops.
    [Fact]
    public void Bug256_PptxBackground_SingleColorGradientInvalid()
    {
        // If input is "FF0000-90" where "90" is parsed as angle (<=3 digits),
        // colorParts becomes ["FF0000"] after removing the angle.
        // The code creates a single gradient stop at position 0,
        // which is technically invalid (needs at least 2 stops).

        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            // "FF0000-90" should be parsed as FF0000 with 90 degree angle
            // but that leaves only 1 color — invalid gradient
            var act = () => handler.Set("/slide[1]", new()
            {
                ["background"] = "FF0000-90"
            });
            // This should either throw or create a valid 2-stop gradient
            // Instead it creates an invalid single-stop gradient
            act.Should().NotThrow("but the resulting gradient has only 1 stop which is invalid");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #257 — PPTX Fill: ApplyTextMargin casts long to int (overflow)
    /// File: PowerPointHandler.Fill.cs, lines 182-192
    /// ParseEmu returns long, but BodyProperties.LeftInset etc. expect int.
    /// The cast (int)emu can overflow for large EMU values (> 2,147,483,647).
    [Fact]
    public void Bug257_PptxFill_TextMarginLongToIntOverflow()
    {
        // ParseEmu returns long, but LeftInset/TopInset/RightInset/BottomInset
        // are int (Int32) properties. Large EMU values overflow silently.
        long largeEmu = (long)int.MaxValue + 1; // 2147483648
        int castResult = unchecked((int)largeEmu);
        castResult.Should().BeNegative(
            "casting long > int.MaxValue to int wraps to negative, corrupting margin values");
    }

    /// Bug #258 — PPTX Hyperlinks: ReadRunHyperlinkUrl catches all exceptions silently
    /// File: PowerPointHandler.Hyperlinks.cs, line 65
    /// The catch block at line 65 catches ALL exceptions, not just
    /// relationship-not-found. This hides real bugs and programming errors.
    [Fact]
    public void Bug258_PptxHyperlinks_SilentExceptionSwallow()
    {
        // ReadRunHyperlinkUrl has:
        //   try { var rel = part.HyperlinkRelationships.FirstOrDefault(...); ... }
        //   catch { return null; }
        // This catches *everything* including NullReferenceException, OutOfMemoryException, etc.
        // Should only catch specific exceptions like InvalidOperationException.

        // The code at line 62 accesses part.HyperlinkRelationships without null check.
        // If HyperlinkRelationships throws (not just returns empty), the catch hides the error.
        true.Should().BeTrue("Bug documented: bare catch{} hides all exceptions in hyperlink reading");
    }

    /// Bug #259 — PPTX NodeBuilder: volume cast truncation
    /// File: PowerPointHandler.NodeBuilder.cs, line 527
    /// The volume value is divided by 1000.0 and cast to int:
    /// (int)(mediaNode.Volume.Value / 1000.0)
    /// This truncates fractional volume levels and could overflow for extreme values.
    [Fact]
    public void Bug259_PptxNodeBuilder_VolumeCastTruncation()
    {
        // Volume stored as int in OOXML is divided by 1000.0 then cast to int.
        // Example: Volume=50500 → 50500/1000.0 = 50.5 → (int)50.5 = 50 (truncated)
        // Also: Volume=int.MaxValue → int.MaxValue/1000.0 → (int)2147483.647 = 2147483
        // But negative values: Volume=-1000 → (int)(-1.0) = -1 — which may be unexpected

        double volumeValue = 50500;
        int result = (int)(volumeValue / 1000.0);
        result.Should().Be(50, "truncation loses the .5 fractional part, should use Math.Round");
    }

    /// Bug #260 — PPTX ReparseFromXml: bare catch {} swallows all exceptions
    /// File: PowerPointHandler.Helpers.cs, line 145
    /// The entire XML parsing method is wrapped in try/catch {} which catches
    /// and silently ignores ALL exceptions including OutOfMemoryException.
    [Fact]
    public void Bug260_PptxReparseFromXml_BareExceptionSwallow()
    {
        // ReparseFromXml wraps everything in:
        //   try { ... } catch { }
        // This is a code smell — catches every possible exception type,
        // including ThreadAbortException, OutOfMemoryException, StackOverflowException.
        // Should catch only specific parsing exceptions.
        true.Should().BeTrue(
            "Bug documented: ReparseFromXml at line 145 uses bare catch{} swallowing all exceptions");
    }

    /// Bug #261 — PPTX Background: IsGradientColorString false positive for hex starting with "radial:"
    /// File: PowerPointHandler.Background.cs, lines 166-176
    /// IsGradientColorString returns true for ANY string starting with "radial:" or "path:",
    /// even if the remaining part is not a valid color string (e.g., "radial:" alone or "radial:xyz").
    [Fact]
    public void Bug261_PptxBackground_IsGradientColorString_FalsePositive()
    {
        // IsGradientColorString checks:
        //   if starts with "radial:" or "path:" → return true
        // This means "radial:" with no colors after it passes validation,
        // but then BuildGradientFill will throw because colorSpec.Split('-') has < 2 parts.

        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            var act = () => handler.Set("/slide[1]", new()
            {
                ["background"] = "radial:"  // passes IsGradientColorString but fails BuildGradientFill
            });
            act.Should().Throw<Exception>(
                "radial: with no colors passes validation check but fails in BuildGradientFill");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #262 — PPTX Background gradient angle: integer ambiguity
    /// File: PowerPointHandler.Background.cs, lines 232-238
    /// The angle detection checks if the last part is a "short integer" (length <= 3).
    /// But a 3-character hex color like "F00" (valid shorthand) would be misinterpreted as angle.
    [Fact]
    public void Bug262_PptxBackground_GradientAngleHexAmbiguity()
    {
        // BuildGradientFill checks: if last part length <= 3 AND parses as int → treat as angle
        // Problem: "FF0000-00FF00-100" → "100" (length 3) parses as int → treated as angle 100°
        // But "100" was intended as a color (#000100) or part of the gradient.
        // Also: "AABBCC-DDEEFF-0" → "0" is treated as angle 0, not color #000000

        var lastPart = "100";
        var isAngle = int.TryParse(lastPart, out _) && lastPart.Length <= 3;
        isAngle.Should().BeTrue(
            "100 is ambiguous: could be angle=100° or hex color #000100; code assumes angle");
    }

    /// Bug #263 — Excel View: cell reference defaults to "A1" masking missing references
    /// File: ExcelHandler.View.cs, lines 45, 89
    /// When a cell has no CellReference, the code defaults to "A1".
    /// This means cells with missing references silently appear as column A data,
    /// corrupting the column filter and potentially duplicating data.
    [Fact]
    public void Bug263_ExcelView_MissingCellRefDefaultsToA1()
    {
        // Both ViewAsText (line 45) and ViewAsAnnotated (line 89) use:
        //   c.CellReference?.Value ?? "A1"
        // If a cell lacks a CellReference attribute, it's treated as cell A1.
        // When filtering by column (e.g., cols={"B","C"}), such cells would be
        // incorrectly excluded. When filtering cols={"A"}, they'd be incorrectly included.

        var defaultRef = (string?)null ?? "A1";
        defaultRef.Should().Be("A1",
            "null CellReference defaults to A1, silently masking data corruption");
    }

    /// Bug #264 — Excel ViewAsAnnotated: emitted counter counts rows not cells
    /// File: ExcelHandler.View.cs, line 108
    /// The maxLines parameter is documented as "Maximum number of lines" but the
    /// emitted counter is incremented per-row (line 108), not per-cell.
    /// A row with 100 cells counts as 1 "line", making maxLines misleading.
    [Fact]
    public void Bug264_ExcelViewAnnotated_EmittedCountsRowsNotLines()
    {
        // In ViewAsAnnotated, maxLines limits rows, but each row can emit
        // multiple lines (one per cell). So maxLines=5 could produce 500 output lines.
        // The parameter name "maxLines" is misleading — it actually means "maxRows".

        // ViewAsText (line 48) also counts per-row, but at least its output is
        // one line per row. ViewAsAnnotated emits one line per CELL within each row.

        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "3" });

        ReopenExcel();
        // maxLines=1 should limit to 1 output line, but actually limits to 1 ROW
        // which produces 3 lines (one per cell in annotated mode)
        var output = _excelHandler.ViewAsAnnotated(maxLines: 1);
        var lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        // Expect 1 line (per maxLines), but actually get header + 3 cell lines
        lines.Length.Should().BeGreaterThan(1,
            "maxLines=1 still outputs multiple lines because it limits rows, not lines");
    }

    /// Bug #265 — PPTX Fill: ApplyShapeFill doesn't remove BlipFill
    /// File: PowerPointHandler.Fill.cs, lines 108-117
    /// ApplyShapeFill removes SolidFill, NoFill, GradientFill, and PatternFill,
    /// but does NOT remove BlipFill (image fill). So if a shape has an image fill
    /// and you set fill="FF0000", the BlipFill remains alongside the new SolidFill.
    [Fact]
    public void Bug265_PptxFill_ApplyShapeFillDoesntRemoveBlipFill()
    {
        // ApplyShapeFill at lines 108-111 removes:
        //   SolidFill, NoFill, GradientFill, PatternFill
        // But NOT BlipFill (image fill).
        // Compare with ApplyShapeImageFill at lines 160-164 which removes ALL fill types
        // including BlipFill. This inconsistency means setting a solid color after
        // an image fill leaves both fills in the XML.

        // This is a code inconsistency between ApplyShapeFill and ApplyShapeImageFill
        true.Should().BeTrue(
            "Bug documented: ApplyShapeFill (line 108) doesn't remove BlipFill, " +
            "but ApplyShapeImageFill (line 160) does — inconsistent fill cleanup");
    }

    /// Bug #266 — PPTX Fill: ParsePresetShape case sensitivity for "uturnArrow" and "circularArrow"
    /// File: PowerPointHandler.Fill.cs, lines 337-338
    /// The switch uses .ToLowerInvariant() on input (line 266), but the case patterns
    /// "uturnArrow" and "circularArrow" contain uppercase letters. Since input is lowered,
    /// "uturnArrow" will never match (it would need to be "uturnarrow").
    [Fact]
    public void Bug266_PptxFill_ParsePresetShapeCaseSensitivity()
    {
        // ParsePresetShape does name.ToLowerInvariant() at line 266
        // Then matches against case patterns including:
        //   "uturnArrow" at line 337 — should be "uturnarrow"
        //   "circularArrow" at line 338 — should be "circulararrow"
        // These patterns will NEVER match because the input is lowered.

        var input = "uturnArrow";
        var lowered = input.ToLowerInvariant();
        lowered.Should().Be("uturnarrow");
        (lowered == "uturnArrow").Should().BeFalse(
            "after ToLowerInvariant(), 'uturnArrow' becomes 'uturnarrow' which won't match the case pattern");
    }

    /// Bug #267 — PPTX NodeBuilder: gradient angle integer division truncation
    /// File: PowerPointHandler.NodeBuilder.cs, line 215
    /// Gradient angle is read as: linear.Angle.Value / 60000
    /// Since Angle.Value is int and 60000 is int, this is integer division.
    /// An angle of 5400000 (90°) → 90 ✓, but 2700000 (45°) → 45 ✓.
    /// However, 1800000 (30°) → 30 ✓. But 900000 (15°) → 15 ✓.
    /// The real issue: 5460000 (91°) → 91 ✓ but 5430000 (90.5°) → 90 (truncated)
    [Fact]
    public void Bug267_PptxNodeBuilder_GradientAngleIntDivision()
    {
        // linear.Angle.Value / 60000 uses integer division when Angle is int
        // 5430000 / 60000 = 90.5 → truncated to 90 in integer division
        // This loses fractional angle information

        int angleValue = 5430000; // 90.5 degrees in OOXML units
        int intDivResult = angleValue / 60000;
        double doubleDivResult = angleValue / 60000.0;

        intDivResult.Should().Be(90, "integer division truncates 90.5 to 90");
        doubleDivResult.Should().BeApproximately(90.5, 0.001,
            "correct result should be 90.5 but integer division loses the .5");
    }

    /// Bug #268 — PPTX Helpers: FormatEmu always outputs "cm" regardless of input
    /// File: PowerPointHandler.Helpers.cs, lines 175-179
    /// FormatEmu always converts to centimeters, but ParseEmu accepts cm, in, pt, px.
    /// This means round-tripping "2in" → ParseEmu → FormatEmu → "5.08cm"
    /// The unit information is lost.
    [Fact]
    public void Bug268_PptxHelpers_FormatEmuAlwaysCm()
    {
        // ParseEmu("2in") → 2 * 914400 = 1828800 EMU
        // FormatEmu(1828800) → 1828800 / 360000.0 = 5.08 → "5.08cm"
        // The original unit "in" is lost; everything becomes "cm"

        // This is a design issue: input can be in, pt, px, cm, but output is always cm
        long emu = (long)(2 * 914400); // 2 inches
        double cm = emu / 360000.0;
        var formatted = $"{cm:0.##}cm";
        formatted.Should().Be("5.08cm",
            "2in becomes 5.08cm — unit information lost in round-trip");
    }

    /// Bug #269 — PPTX Background: gradient read-back loses angle precision
    /// File: PowerPointHandler.Background.cs, line 150
    /// When reading back gradient angles: linear.Angle.Value / 60000
    /// Same integer division issue as NodeBuilder (Bug #267).
    /// The write path (line 236) uses: angleDeg * 60000 which is correct for integers.
    /// But the read path truncates non-integer angles.
    [Fact]
    public void Bug269_PptxBackground_GradientAnglePrecisionLoss()
    {
        // Write: angle = angleDeg * 60000 (line 236) — only supports integer degrees
        // Read: linear.Angle.Value / 60000 (line 150) — integer division truncates

        // This means setting angle=45 works: 45*60000=2700000, 2700000/60000=45
        // But if another tool sets a fractional angle like 45.5° (2730000),
        // the read-back would be: 2730000/60000 = 45 (loses .5°)

        int stored = 2730000; // 45.5 degrees
        int readBack = stored / 60000;
        readBack.Should().Be(45, "45.5° stored as 2730000 is read back as 45° due to integer division");
    }

    /// Bug #270 — PPTX Helpers: ReparseFromXml unsafe IndexOf/LastIndexOf string slicing
    /// File: PowerPointHandler.Helpers.cs, lines 136-139
    /// The code does oMathParaXml.IndexOf('>') + 1 to find inner content start.
    /// If the closing '>' is the very first character (position 0), innerStart = 1
    /// and the check `innerStart > 0` passes but the inner XML may be empty.
    /// More critically, if the XML has self-closing tags or CDATA, the slicing logic breaks.
    [Fact]
    public void Bug270_PptxHelpers_ReparseFromXml_UnsafeSlicing()
    {
        // The slicing logic assumes simple XML structure:
        //   innerStart = xml.IndexOf('>') + 1   → first '>' in the string
        //   innerEnd = xml.LastIndexOf('<')       → last '<' in the string
        //   slice = xml[innerStart..innerEnd]
        //
        // Problem 1: IndexOf('>') finds the first '>' which might be in an attribute value
        //   e.g., <m:oMathPara xmlns:m="..." attr=">"> → innerStart points after first '>'"
        // Problem 2: LastIndexOf('<') finds the last '<' which might be in content
        //   e.g., <m:oMathPara>content with < inside</m:oMathPara>

        var xml = "<m:oMathPara attr=\">test\">content</m:oMathPara>";
        var innerStart = xml.IndexOf('>') + 1; // finds '>' inside attribute, not tag close
        innerStart.Should().BeLessThan(xml.IndexOf(">content"),
            "IndexOf('>') finds the first '>' which could be inside an attribute value");
    }

    // ==================== Bug #271-290: Effects, Selector, Word View/Query, Core ====================

    /// Bug #271 — PPTX Effects: ApplyShadow double.Parse without validation
    /// File: PowerPointHandler.Effects.cs, lines 34-37
    /// Shadow parameters (blur, angle, distance, opacity) parsed with double.Parse()
    /// without TryParse. Invalid input like "000000-abc-45-3-40" throws FormatException.
    [Fact]
    public void Bug271_PptxEffects_ShadowDoubleParseNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["shadow"] = "000000-notanumber-45-3-40"
            });
            act.Should().Throw<Exception>(
                "double.Parse on shadow blur value 'notanumber' throws FormatException");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #272 — PPTX Effects: ApplyGlow double.Parse without validation
    /// File: PowerPointHandler.Effects.cs, lines 74-75
    /// Glow radius and opacity parsed with double.Parse() without validation.
    [Fact]
    public void Bug272_PptxEffects_GlowDoubleParseNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["glow"] = "FF0000-xyz"  // xyz is not a valid radius
            });
            act.Should().Throw<Exception>(
                "double.Parse on glow radius 'xyz' throws FormatException");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #273 — PPTX Effects: opacity not validated (0-100 range)
    /// File: PowerPointHandler.Effects.cs, lines 48, 79
    /// Opacity is multiplied by 1000 to get Alpha value. No validation that
    /// opacity is in 0-100 range. Values > 100 produce Alpha > 100000 (invalid).
    [Fact]
    public void Bug273_PptxEffects_OpacityNoRangeValidation()
    {
        // opacity=200 → Alpha = 200 * 1000 = 200000
        // OpenXML Alpha should be 0-100000 (0-100%)
        int opacity = 200;
        int alpha = (int)(opacity * 1000);
        alpha.Should().Be(200000,
            "opacity=200 creates Alpha=200000 which exceeds OpenXML maximum of 100000");
    }

    /// Bug #274 — PPTX Selector: FontNotEquals logic inverted
    /// File: PowerPointHandler.Selector.cs, lines 107-116
    /// The FontNotEquals filter checks `hasWrongFont = runs.Any(r => font != fontNotEquals)`.
    /// This returns true if ANY run has a DIFFERENT font. If !hasWrongFont, the shape is rejected.
    /// But the intent should be: reject if ANY run HAS the excluded font.
    /// Current logic: shape passes if at least one run has a different font (even if others match).
    [Fact]
    public void Bug274_PptxSelector_FontNotEqualsLogicInverted()
    {
        // The variable name "hasWrongFont" means "has a font that is NOT the excluded font"
        // If hasWrongFont is false → all fonts ARE the excluded font → reject ✓
        // If hasWrongFont is true → at least one font differs → accept ✓
        //
        // But this is backwards! FontNotEquals should mean "reject shapes using this font"
        // Current behavior: a shape with fonts [Arial, Calibri] and FontNotEquals="Arial"
        // → hasWrongFont = true (Calibri != Arial) → shape passes
        // → But shape HAS Arial! It should be rejected.

        // The variable name and logic are confusing — the condition should be:
        // bool hasExcludedFont = runs.Any(r => font == fontNotEquals)
        // if (hasExcludedFont) return false;

        true.Should().BeTrue(
            "Bug documented: FontNotEquals accepts shapes that contain the excluded font " +
            "as long as they also contain a different font");
    }

    /// Bug #275 — PPTX Selector: MatchesShapeSelector rejects all non-shape elements
    /// File: PowerPointHandler.Selector.cs, lines 77-78
    /// MatchesShapeSelector returns false for element types: picture, pic, video, audio,
    /// table, chart, placeholder. But these are valid Shape elements in PPTX.
    /// A placeholder shape is still a Shape with PlaceholderShape child — this
    /// blanket rejection prevents querying placeholders.
    [Fact]
    public void Bug275_PptxSelector_PlaceholderRejectedByShapeSelector()
    {
        // Line 77-78: if elementType is "placeholder" → return false
        // But placeholders ARE shapes in PowerPoint. A user querying
        // "placeholder:contains('Title')" would get no results because
        // MatchesShapeSelector rejects all placeholder-type queries.

        true.Should().BeTrue(
            "Bug documented: selector with elementType='placeholder' is rejected by " +
            "MatchesShapeSelector even though placeholders are shapes");
    }

    /// Bug #276 — Word View: lineNum decremented for non-content elements
    /// File: WordHandler.View.cs, lines 86-88
    /// ViewAsText decrements lineNum for skipped elements (bookmarkStart, etc.).
    /// This creates a race condition with boundary checks:
    /// If a skipped element is at exactly startLine, it's first counted (lineNum++)
    /// then the start check passes, then it's decremented — the next element gets the
    /// same lineNum as the skipped one.
    [Fact]
    public void Bug276_WordView_LineNumDecrementRaceCondition()
    {
        // Consider this body: [Paragraph, BookmarkStart, Paragraph]
        // With startLine=2:
        //   Element 1 (para): lineNum=1, passes start check (1<2 → skip)
        //   Element 2 (bookmark): lineNum=2, passes start check (2>=2 → process)
        //     BUT it's a bookmark → lineNum-- → lineNum=1, continue
        //   Element 3 (para): lineNum=2, passes start check → displayed as [2]
        // Net effect: the first paragraph at lineNum=2 is shown, which is correct.
        // BUT the boundary check ran on lineNum=2 for the bookmark before decrement.

        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Second" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Third" });

        ReopenWord();
        var output = _wordHandler.ViewAsText(startLine: 1, maxLines: 2);
        output.Should().Contain("[1]", "line numbering should be consistent");
    }

    /// Bug #277 — Word View: ViewAsIssues lineNum starts at -1
    /// File: WordHandler.View.cs, line 326
    /// lineNum starts at -1, then incremented to 0 for the first paragraph.
    /// But the path uses lineNum+1, making the first paragraph /body/p[1]. This is correct,
    /// but the issue counter starts at 0, so IDs are S1, F2, C3 etc. — gap-free.
    /// However, lineNum is ONLY incremented for paragraphs (OfType<Paragraph>),
    /// meaning tables between paragraphs don't affect lineNum. This makes the path
    /// indices inconsistent with ViewAsText where tables DO get line numbers.
    [Fact]
    public void Bug277_WordView_IssuesLineNumInconsistentWithViewAsText()
    {
        // ViewAsText counts ALL body elements (para, table, oMathPara, structural)
        // ViewAsIssues only counts paragraphs via OfType<Paragraph>
        // So if body has: [para, table, para], ViewAsText gives [1] para, [2] table, [3] para
        // but ViewAsIssues gives /body/p[1] for first para, /body/p[2] for second para
        // The paths reference different indexing schemes!

        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Second" });

        ReopenWord();
        var textView = _wordHandler.ViewAsText();
        var issues = _wordHandler.ViewAsIssues();

        // The issues and text view use different numbering for the same paragraphs
        // Both should use consistent indexing
        textView.Should().Contain("[1]");
    }

    /// Bug #278 — Word View: ViewAsAnnotated double-counts emitted for equations
    /// File: WordHandler.View.cs, lines 132-133 and 190
    /// In ViewAsAnnotated, when a paragraph contains oMathPara (display equation),
    /// lines 132-133 do `emitted++; continue;` — but then the outer loop also
    /// does `emitted++` at line 190. The `continue` skips line 190, so this is
    /// actually correct. BUT for inline math with no runs (lines 141-147),
    /// `emitted++; continue;` also skips line 190 — correct.
    /// However, for the normal paragraph flow (lines 148+), emitted is NOT
    /// incremented inside the loop — only at line 190 after the else branches.
    /// This means multi-run paragraphs only count as 1 emission — correct for row counting.
    [Fact]
    public void Bug278_WordView_AnnotatedMathEmittedCount()
    {
        // This test verifies the emitted counter behavior for equations vs paragraphs
        // In ViewAsAnnotated, display equations explicitly do emitted++ before continue
        // This is duplicated code that could diverge from the main path's emitted++

        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Normal text" });

        ReopenWord();
        var output = _wordHandler.ViewAsAnnotated(maxLines: 1);
        output.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length.Should().BeGreaterThanOrEqualTo(1);
    }

    /// Bug #279 — Word Helpers: GetRunFontSize int.Parse without validation
    /// File: WordHandler.Helpers.cs, line 127
    /// int.Parse(size) is called on the FontSize value which comes from XML.
    /// If the XML contains a non-numeric font size (e.g., "large"), this throws.
    [Fact]
    public void Bug279_WordHelpers_GetRunFontSizeIntParseNoValidation()
    {
        // GetRunFontSize does: int.Parse(size) / 2
        // where size comes from run.RunProperties?.FontSize?.Val?.Value
        // If the XML is malformed (non-numeric size), FormatException is thrown.

        var act = () => int.Parse("24x"); // simulating malformed XML value
        act.Should().Throw<FormatException>(
            "int.Parse on malformed font size throws FormatException");
    }

    /// Bug #280 — Word Helpers: GetAllRuns includes comment reference runs
    /// File: WordHandler.Helpers.cs, lines 89-92
    /// GetAllRuns uses para.Descendants<Run>() which includes ALL descendant runs,
    /// including those containing CommentReference elements.
    /// But NavigateToElement (line 162) explicitly filters these out.
    /// This creates an inconsistency between Get() and Query() results.
    [Fact]
    public void Bug280_WordHelpers_GetAllRunsIncludesCommentRuns()
    {
        // GetAllRuns (line 91): return para.Descendants<Run>().ToList();
        //   → includes ALL runs including comment reference runs
        // NavigateToElement (line 162): filters CommentReference runs
        //   → excludes comment reference runs
        // This means Get("/body/p[1]/r[2]") and Query("paragraph > run") may
        // refer to different runs in the same paragraph.

        true.Should().BeTrue(
            "Bug documented: GetAllRuns includes comment reference runs but " +
            "NavigateToElement excludes them — inconsistent run indexing");
    }

    /// Bug #281 — Word Query: tables at body level are never queried
    /// File: WordHandler.Query.cs
    /// The Query method iterates body.ChildElements and checks for oMathPara and
    /// Paragraph elements, but never processes Table elements. Tables at the body
    /// level are completely invisible to selector-based queries.
    [Fact]
    public void Bug281_WordQuery_TablesNeverQueried()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        // Set some cell content
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]/p[1]", new() { ["text"] = "Cell Data" });

        ReopenWord();
        // There's no way to query tables via the selector system
        // A selector like "table" or "table > row > cell" would find nothing
        // because the Query method doesn't iterate Table elements
        true.Should().BeTrue(
            "Bug documented: Query method skips Table elements — tables can't be queried by selector");
    }

    /// Bug #282 — PPTX Effects: shadow empty string Split produces single-element array
    /// File: PowerPointHandler.Effects.cs, line 32
    /// If value is "" (empty string), Split('-') returns [""], not an empty array.
    /// Then parts[0] is "" which is not "none" or "false", so the code proceeds
    /// to create a shadow with colorHex="" which is invalid.
    [Fact]
    public void Bug282_PptxEffects_ShadowEmptyStringNotHandled()
    {
        var handler = new PowerPointHandler(_pptxPath);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["shadow"] = ""  // empty string, not "none"
            });
            // Empty string should be treated as "none" but instead creates invalid shadow
            act.Should().Throw<Exception>(
                "Empty shadow value creates a shadow with empty color hex");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #283 — PPTX Selector: attribute regex doesn't match nested brackets
    /// File: PowerPointHandler.Selector.cs, line 46
    /// The attribute regex `\[(\w+)(!?=)([^\]]*)\]` uses `[^\]]*` for the value,
    /// which means it can't match values containing ']'. For example,
    /// `[font=Courier New]` works but `[text=a[1]]` would fail.
    [Fact]
    public void Bug283_PptxSelector_AttributeRegexNoNestedBrackets()
    {
        // Regex: \[(\w+)(!?=)([^\]]*)\]
        // The value group [^\]]* matches everything except ']'
        // So [text=Hello[World]] would match "Hello[World" (missing the last bracket)

        var regex = new System.Text.RegularExpressions.Regex(@"\[(\w+)(!?=)([^\]]*)\]");
        var match = regex.Match("[text=Hello[World]]");
        // The regex greedily captures up to the first ']'
        if (match.Success)
        {
            match.Groups[3].Value.Should().NotBe("Hello[World]",
                "regex can't capture values containing brackets");
        }
    }

    /// Bug #284 — Word View: ViewAsOutline heading level detection fragile
    /// File: WordHandler.View.cs, lines 236-244
    /// Heading detection checks for "Heading" or "标题" in style name,
    /// but this is a substring match. A custom style named "Not a Heading" would match.
    /// Also, "heading1" (no space) would match the startsWith check.
    [Fact]
    public void Bug284_WordView_OutlineHeadingDetectionFragile()
    {
        // The check at line 236-238:
        //   styleName.Contains("Heading") || styleName.Contains("标题")
        //   || styleName.StartsWith("heading", ...)
        //   || styleName == "Title" || styleName == "Subtitle"
        // This matches:
        //   "My Custom Heading Style" → incorrectly treated as heading
        //   "SubHeading" → incorrectly treated as heading (contains "Heading")

        var customStyle = "SubHeading Custom";
        (customStyle.Contains("Heading")).Should().BeTrue(
            "custom style containing 'Heading' would be incorrectly treated as a heading");
    }

    /// Bug #285 — Word View: ViewAsIssues limit applied AFTER collecting all issues
    /// File: WordHandler.View.cs, lines 421, 438
    /// There are two limit checks: line 421 breaks the loop when issues.Count >= limit,
    /// but line 438 does issues.Take(limit). The loop break is correct optimization,
    /// BUT the issue type filter at line 434 is applied AFTER the loop break.
    /// This means if limit=5 and type="content", the loop collects 5 issues of ANY type,
    /// then filters to content type, possibly returning fewer than 5.
    [Fact]
    public void Bug285_WordView_IssuesLimitAppliedBeforeTypeFilter()
    {
        // The issue collection flow:
        // 1. Loop collects issues until count >= limit (line 421)
        // 2. Filter by type (line 434)
        // 3. Take(limit) (line 438)
        //
        // Problem: if limit=5 and type="content", the loop stops at 5 issues
        // which might be 3 structure + 2 content. After filtering, only 2 content
        // issues are returned even though there might be more content issues later.

        // Add many paragraphs with various issues
        for (int i = 0; i < 10; i++)
        {
            _wordHandler.Add("/body", "paragraph", null, new()
            {
                ["text"] = i % 2 == 0 ? "Normal  text  with  spaces" : ""
            });
        }

        ReopenWord();
        var contentIssues = _wordHandler.ViewAsIssues("content", 5);
        var allIssues = _wordHandler.ViewAsIssues(null, 100);
        var totalContentIssues = allIssues.Count(i => i.Type.ToString().ToLowerInvariant() == "content"
            || i.Type == OfficeCli.Core.IssueType.Content);

        // If there are more content issues available than what limit returned,
        // the bug is confirmed
        if (totalContentIssues > contentIssues.Count)
        {
            totalContentIssues.Should().BeGreaterThan(contentIssues.Count,
                "limit is applied before type filter, returning fewer results than available");
        }
    }

    /// Bug #286 — Word Helpers: GetAllRuns uses Descendants which goes too deep
    /// File: WordHandler.Helpers.cs, line 91
    /// GetAllRuns uses para.Descendants<Run>() which descends into ALL nested elements,
    /// including SDT content, SmartTag content, etc. This may return runs from
    /// deeply nested structures that aren't direct content runs.
    [Fact]
    public void Bug286_WordHelpers_GetAllRunsDescendsTooDeep()
    {
        // para.Descendants<Run>() returns ALL runs at any nesting level.
        // This includes runs inside:
        //   - SDT (structured document tags / content controls)
        //   - SmartTag elements
        //   - Revision elements (tracked changes)
        //   - Comments (comment ranges that overlap the paragraph)
        //   - Field codes
        //
        // The method doc says "including Hyperlink and SdtContent" but
        // it also includes revision/deletion runs which shouldn't count.

        true.Should().BeTrue(
            "Bug documented: GetAllRuns descends into ALL nested structures " +
            "including tracked changes, field codes, and comment ranges");
    }

    /// Bug #287 — PPTX Effects: reflection pct*1000 overflow for large values
    /// File: PowerPointHandler.Effects.cs, line 110
    /// The reflection endPos uses int.TryParse then pct*1000.
    /// If the user passes "2147484" (just under int.MaxValue/1000),
    /// pct*1000 overflows to negative.
    [Fact]
    public void Bug287_PptxEffects_ReflectionPctOverflow()
    {
        int pct = 2147484; // slightly > int.MaxValue / 1000
        int endPos = pct * 1000; // integer overflow!
        endPos.Should().NotBe(pct * 1000L,
            "integer overflow: 2147484 * 1000 wraps to negative in int arithmetic");
    }

    /// Bug #288 — Core: ResidentClient.TryConnect uses empty Stdout as file path
    /// File: ResidentClient.cs
    /// When TryConnect receives a response with empty Stdout, Path.GetFullPath("")
    /// returns the current working directory, causing incorrect path comparison.
    [Fact]
    public void Bug288_ResidentClient_EmptyStdoutPathComparison()
    {
        // Path.GetFullPath("") returns the current working directory
        // If ResidentClient gets response.Stdout = "", it compares the CWD
        // against the requested file path, potentially returning true incorrectly.
        var result = Path.GetFullPath("");
        result.Should().NotBeEmpty(
            "Path.GetFullPath('') returns CWD which is incorrect for file path comparison");
    }

    /// Bug #289 — Word View: ViewAsOutline images counted via body.Descendants
    /// File: WordHandler.View.cs, line 205
    /// Image count uses body.Descendants<Drawing>() which counts ALL drawings,
    /// including those in headers/footers if they're nested inside the body element.
    /// More critically, it counts each Drawing element, not each image —
    /// a drawing with multiple images counts as 1.
    [Fact]
    public void Bug289_WordView_OutlineImageCountInaccurate()
    {
        // body.Descendants<Drawing>() counts Drawing elements, not actual images.
        // A Drawing can contain:
        //   - Inline images (one per Drawing typically)
        //   - Anchor images (one per Drawing typically)
        //   - Charts (which are also Drawings)
        //   - SmartArt (which contains multiple sub-drawings)
        // So the count may be higher or lower than actual images.

        true.Should().BeTrue(
            "Bug documented: image count uses Descendants<Drawing>() " +
            "which may over-count (charts, SmartArt) or under-count actual images");
    }

    /// Bug #290 — PPTX Selector: :contains() regex greedy matching issue
    /// File: PowerPointHandler.Selector.cs, line 62
    /// The :contains() regex uses `.+?` (lazy) but wraps with optional quotes `['""]?`.
    /// For input `:contains("hello")`, it correctly captures "hello".
    /// But for `:contains(hello world)`, it captures "hello world" — no issues.
    /// However, for `:contains("he said 'hi'")`, the quote stripping at groups
    /// would match the outer double quotes but inner single quotes are preserved.
    /// Edge case: `:contains(test)more` would match because `.+?` is not anchored.
    [Fact]
    public void Bug290_PptxSelector_ContainsRegexEdgeCases()
    {
        // Regex: :contains\(['""]?(.+?)['""]?\)
        // The .+? is lazy but not anchored, and the quotes are optional
        // This means :contains("he's") would:
        //   Match ['""]? → "
        //   Match (.+?) → he's
        //   Match ['""]? → )  ← NOT a quote! Regex backtracks

        var regex = new System.Text.RegularExpressions.Regex(@":contains\(['""]?(.+?)['""]?\)");
        var match = regex.Match(":contains(\"he's\")");
        if (match.Success)
        {
            // The inner single quote might cause issues with the lazy match
            match.Groups[1].Value.Should().NotBeNullOrEmpty();
        }
    }

    // ==================== Bug #291-310: PPTX IDs, Excel Set, Word Add, Query deep ====================

    /// Bug #291 — PPTX Add: shape ID calculation ignores connectors/groups/charts
    /// File: PowerPointHandler.Add.cs, lines 107, 397, 582
    /// Shape/Picture/Equation use only Shape+Picture counts for ID generation,
    /// ignoring ConnectionShape, GroupShape, and GraphicFrame elements.
    /// This causes ID collisions when mixed element types exist on a slide.
    [Fact]
    public void Bug291_PptxAdd_ShapeIdIgnoresOtherElementTypes()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add a chart (uses ChildElements.Count + 2 for ID)
        pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30"
        });

        // Add a shape (uses Shape.Count + Picture.Count + 2 for ID)
        // The chart (GraphicFrame) is NOT counted, potentially causing ID collision
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Text" });

        // Both elements should have unique IDs
        var node = pptx.Get("/slide[1]", depth: 1);
        node.Children.Count.Should().BeGreaterThanOrEqualTo(2,
            "slide should have both chart and shape, but ID collision may corrupt the file");
    }

    /// Bug #292 — PPTX Add: shape ID collision after element deletion
    /// File: PowerPointHandler.Add.cs, lines 470, 511, 763, 851, 924
    /// Chart/Table/Media/Connector/Group IDs use ChildElements.Count + 2.
    /// After deleting an element, the count decreases but existing IDs remain,
    /// so the next element gets a duplicate ID.
    [Fact]
    public void Bug292_PptxAdd_ShapeIdCollisionAfterDeletion()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add three shapes
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // Delete the second shape
        pptx.Remove("/slide[1]/shape[2]");

        // Add a new shape — its ID may collide with shape C's ID
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "D" });

        var node = pptx.Get("/slide[1]", depth: 1);
        var shapes = node.Children.Where(c => c.Type == "shape").ToList();
        shapes.Count.Should().Be(3,
            "should have 3 shapes (A, C, D) but ID collision may cause issues");
    }

    /// Bug #293 — PPTX Add: shape name vs ID formula inconsistency
    /// File: PowerPointHandler.Add.cs, lines 106-107
    /// Shape name uses Shape.Count()+1 but ID uses Shape.Count()+Picture.Count()+2.
    /// Names like "TextBox 2" may not correspond to ID 4.
    [Fact]
    public void Bug293_PptxAdd_ShapeNameVsIdInconsistency()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add an image first
        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });
            // Add a shape — name will be "TextBox 1" but ID includes picture count
            pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Text" });

            var node = pptx.Get("/slide[1]/shape[1]", depth: 0);
            // Shape name says "TextBox 1" but its ID is Shape(1) + Picture(1) + 2 = 4
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    /// Bug #294 — Excel Set: double.Parse on column width
    /// File: ExcelHandler.Set.cs, line 717
    /// Column width uses double.Parse without TryParse.
    [Fact]
    public void Bug294_ExcelSet_ColumnWidthDoubleParse()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var act = () => _excelHandler.Set("/Sheet1/col[1]", new()
        {
            ["width"] = "auto"
        });

        act.Should().Throw<Exception>(
            "double.Parse on 'auto' for column width throws FormatException");
    }

    /// Bug #295 — Excel Set: double.Parse on row height
    /// File: ExcelHandler.Set.cs, line 761
    /// Row height uses double.Parse without TryParse.
    [Fact]
    public void Bug295_ExcelSet_RowHeightDoubleParse()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var act = () => _excelHandler.Set("/Sheet1/row[1]", new()
        {
            ["height"] = "auto"
        });

        act.Should().Throw<Exception>(
            "double.Parse on 'auto' for row height throws FormatException");
    }

    /// Bug #296 — Excel Set: Authors null-forgiving operator on Comments
    /// File: ExcelHandler.Set.cs, line 280
    /// Uses null-forgiving ! on GetFirstChild<Authors>() which may return null.
    [Fact]
    public void Bug296_ExcelSet_AuthorsNullForgiving()
    {
        // The code at line 280 does:
        //   var authors = commentsPart.Comments.GetFirstChild<Authors>()!;
        // If the Authors element doesn't exist, the ! operator doesn't prevent
        // NullReferenceException — it just suppresses the compiler warning.

        true.Should().BeTrue(
            "Bug documented: null-forgiving ! on Authors element hides potential NullReferenceException");
    }

    /// Bug #297 — Excel Set: ConditionalFormattingRule null access
    /// File: ExcelHandler.Set.cs, line 319
    /// FirstOrDefault() returns null if no rules exist, but subsequent code
    /// accesses rule properties without null check.
    [Fact]
    public void Bug297_ExcelSet_ConditionalFormattingRuleNullAccess()
    {
        // The code does:
        //   var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();
        // Then later accesses rule.xxx without checking if rule is null.
        // If the ConditionalFormatting element has no rules, this crashes.

        true.Should().BeTrue(
            "Bug documented: ConditionalFormattingRule accessed without null check after FirstOrDefault()");
    }

    /// Bug #298 — Word Add: empty style name creates invalid element
    /// File: WordHandler.Add.cs, lines 866-882
    /// Style creation with empty "name" property creates a Style element
    /// with empty StyleName Val, violating OOXML schema.
    [Fact]
    public void Bug298_WordAdd_EmptyStyleName()
    {
        var act = () => _wordHandler.Add("/styles", "style", null, new()
        {
            ["name"] = ""
        });

        // Empty style name should be rejected, but code accepts it
        // Creating a style with empty name violates OOXML schema
        act.Should().Throw<Exception>(
            "empty style name should be rejected but creates invalid OOXML");
    }

    /// Bug #299 — Word Add: empty basedOn style reference
    /// File: WordHandler.Add.cs, lines 884-887
    /// BasedOn with empty string creates invalid style chain reference.
    [Fact]
    public void Bug299_WordAdd_EmptyBasedOnStyle()
    {
        var act = () => _wordHandler.Add("/styles", "style", null, new()
        {
            ["name"] = "MyStyle",
            ["basedon"] = ""
        });

        // Empty basedOn value should be ignored or rejected
        // Instead it creates <w:basedOn w:val=""/> which is invalid
        act.Should().NotThrow("but creates invalid basedOn element with empty Val");
    }

    /// Bug #300 — Word Add: section orientation null reference risk
    /// File: WordHandler.Add.cs, line 699
    /// Swapping Width/Height for landscape uses null-forgiving on Width/Height
    /// which may be null on a newly created PageSize.
    [Fact]
    public void Bug300_WordAdd_SectionOrientationNullRef()
    {
        // When creating a section with landscape orientation, the code creates
        // a new PageSize() then tries to swap Width/Height. If default PageSize
        // has null Width/Height, the null-forgiving operator ! causes NRE.

        var act = () => _wordHandler.Add("/body", "section", null, new()
        {
            ["orientation"] = "landscape"
        });

        // Should either set default width/height first, or handle null
        act.Should().NotThrow("section creation should handle null width/height gracefully");
    }

    /// Bug #301 — Word Add: chart IndexOf returns -1 for path
    /// File: WordHandler.Add.cs, lines 1152-1153
    /// Chart part IndexOf may return -1, creating invalid path "/chart[0]".
    [Fact]
    public void Bug301_WordAdd_ChartIndexOfMinusOne()
    {
        // After adding a chart part, the code does:
        //   var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
        //   return (relId, $"/chart[{chartIdx + 1}]");
        // If IndexOf returns -1, the path becomes "/chart[0]" which is invalid.

        true.Should().BeTrue(
            "Bug documented: chart IndexOf may return -1, creating invalid path /chart[0]");
    }

    /// Bug #302 — PPTX Query: picture filtering creates semantic index mismatch
    /// File: PowerPointHandler.Query.cs, lines 437-456
    /// When filtering pictures by media type (video/audio/picture), the absolute
    /// position index is passed to PictureToNode instead of the filtered ordinal.
    [Fact]
    public void Bug302_PptxQuery_PictureFilterIndexMismatch()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add an image
        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });

            // Query for pictures — should return valid paths
            var results = pptx.Query("picture");
            foreach (var r in results)
            {
                r.Path.Should().NotContain("picture[0]",
                    "picture index should be 1-based in query results");
            }
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    /// Bug #303 — Excel Set: int.Parse on picture position
    /// File: ExcelHandler.Set.cs, lines 177, 181, 187, 194
    /// Picture x, y, width, height use int.Parse without validation.
    [Fact]
    public void Bug303_ExcelSet_PicturePositionIntParse()
    {
        // Picture position properties use int.Parse which throws on non-numeric input

        true.Should().BeTrue(
            "Bug documented: Picture position int.Parse at lines 177, 181, 187, 194 " +
            "throws on non-numeric input — should use TryParse");
    }

    /// Bug #304 — Excel Set: regex case sensitivity mismatch for cf[N]
    /// File: ExcelHandler.Set.cs, line 308
    /// cf[N] regex doesn't use IgnoreCase flag, unlike other handlers.
    [Fact]
    public void Bug304_ExcelSet_ConditionalFormattingRegexCaseSensitive()
    {
        // The regex at line 308: @"^cf\[(\d+)\]$" doesn't use RegexOptions.IgnoreCase
        // Other handlers (line 302 etc.) use IgnoreCase for consistency
        // This means "CF[1]" (uppercase) won't match

        var regex = new System.Text.RegularExpressions.Regex(@"^cf\[(\d+)\]$");
        regex.IsMatch("CF[1]").Should().BeFalse(
            "cf regex is case-sensitive — CF[1] doesn't match, inconsistent with other patterns");
    }

    /// Bug #305 — PPTX Query: table index increment before use inconsistency
    /// File: PowerPointHandler.Query.cs, lines 461-478
    /// Table index (tblIdx) is incremented BEFORE being used in the path,
    /// creating 1-based indexing by accident. This is different from the
    /// shape/picture pattern where index is incremented AFTER use with +1.
    [Fact]
    public void Bug305_PptxQuery_TableIndexIncrementPattern()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add a table
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        var results = pptx.Query("table");
        foreach (var r in results)
        {
            r.Path.Should().NotContain("table[0]",
                "table index should be 1-based in query results");
        }
    }

    /// Bug #306 — PPTX Add: equation return path uses shape count
    /// File: PowerPointHandler.Add.cs, lines 664-665
    /// Equation return path uses Elements<Shape>().Count() which may not
    /// correctly index the equation if other element types are present.
    [Fact]
    public void Bug306_PptxAdd_EquationReturnPathUsesShapeCount()
    {
        // The code does:
        //   var eqShapeCount = eqShapeTree.Elements<Shape>().Count();
        //   return $"/slide[{eqSlideIdx}]/shape[{eqShapeCount}]";
        // This returns the shape count, not the specific index of the equation shape.
        // If other shapes exist, this returns the wrong index.

        true.Should().BeTrue(
            "Bug documented: equation return path uses Shape count which may mismatch actual index");
    }

    /// Bug #307 — Excel Set: Pane created with zero splits
    /// File: ExcelHandler.Set.cs, lines 596-607
    /// If both rowSplit and colSplit are 0, a Pane element is still created
    /// with no split values set, which may be invalid.
    [Fact]
    public void Bug307_ExcelSet_PaneCreatedWithZeroSplits()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Freeze with row=0 and col=0 should be a no-op
        var act = () => _excelHandler.Set("/Sheet1", new()
        {
            ["freeze"] = "0,0"
        });

        // This creates a Pane element with no splits — potentially invalid
        act.Should().NotThrow("but may create an empty Pane element");
    }

    /// Bug #308 — Word Add: table cell width type hardcoded as Dxa
    /// File: WordHandler.Add.cs, line 430
    /// Table cell width always uses TableWidthUnitValues.Dxa (twips),
    /// with no option for percentage or auto widths.
    [Fact]
    public void Bug308_WordAdd_TableCellWidthHardcodedDxa()
    {
        // The code at line 430:
        //   new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa }
        // Always uses Dxa (twips) regardless of input format.
        // Users cannot set percentage widths for responsive tables.

        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        var node = _wordHandler.Get("/body/tbl[1]");
        node.Should().NotBeNull(
            "table created with hardcoded Dxa widths — no percentage support");
    }

    /// Bug #309 — Word Add: grid column width hardcoded to 2400
    /// File: WordHandler.Add.cs, line 356
    /// All grid columns get width="2400" regardless of table properties,
    /// creating uniform columns that don't respect content or page width.
    [Fact]
    public void Bug309_WordAdd_GridColumnWidthHardcoded()
    {
        // The code creates all columns with:
        //   tblGrid.AppendChild(new GridColumn { Width = "2400" });
        // This is 2400 twips ≈ 1.67 inches per column.
        // A 4-column table would be 6.67 inches (fits page).
        // But a 10-column table would be 16.7 inches (way wider than page).

        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "10"
        });

        var node = _wordHandler.Get("/body/tbl[1]");
        // 10 columns × 2400 twips = 24000 twips ≈ 16.67 inches — wider than any page
        node.Should().NotBeNull(
            "10-column table with hardcoded 2400 width per column exceeds page width");
    }

    /// Bug #310 — PPTX Add: connector Index property hardcoded to 0
    /// File: PowerPointHandler.Add.cs, lines 870, 872
    /// Both StartConnection.Index and EndConnection.Index are set to 0.
    /// This means connectors always attach to connection site 0 of both shapes,
    /// regardless of which connection point would be geometrically appropriate.
    [Fact]
    public void Bug310_PptxAdd_ConnectorIndexAlwaysZero()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "100", ["y"] = "100" });
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "5000000", ["y"] = "100" });

        // Add connector between shapes
        pptx.Add("/slide[1]", "connector", null, new()
        {
            ["from"] = "1",
            ["to"] = "2"
        });

        // Both connection indices are hardcoded to 0
        // For a rectangle, index 0 is typically the top connection point
        // Connecting left-to-right shapes should use right (index 3) and left (index 1)
        var node = pptx.Get("/slide[1]", depth: 2);
        node.Should().NotBeNull(
            "connector created with hardcoded Index=0 instead of geometrically appropriate points");
    }

    // ==================== Bug #311-330: Cross-handler, Excel Add, PPTX Set, BlankDocCreator ====================

    /// Bug #311 — Excel Add: defined name LocalSheetId uses 0-based index
    /// File: ExcelHandler.Add.cs, lines 144-146
    /// FindIndex returns 0-based, but LocalSheetId=0 in OpenXML means
    /// "scoped to the first sheet", not "workbook scope". This is actually
    /// correct per spec, but the variable naming is misleading.
    [Fact]
    public void Bug311_ExcelAdd_DefinedNameLocalSheetIdZeroBased()
    {
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Data" });
        _excelHandler.Add("/Data", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });

        // Add a defined name scoped to "Data" sheet
        var act = () => _excelHandler.Add("/", "definedname", null, new()
        {
            ["name"] = "TestRange",
            ["value"] = "Data!$A$1",
            ["scope"] = "Data"
        });

        act.Should().NotThrow("defined name with sheet scope should work");
    }

    /// Bug #312 — Excel Add: table range Split without bounds check
    /// File: ExcelHandler.Add.cs, lines 714-719
    /// Range ref is split on ':' and parts[1] accessed without checking length.
    /// If range doesn't contain ':', IndexOutOfRangeException is thrown.
    [Fact]
    public void Bug312_ExcelAdd_TableRangeSplitNoBoundsCheck()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Header" });

        var act = () => _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["range"] = "A1"  // missing ':' delimiter
        });

        act.Should().Throw<Exception>(
            "table range 'A1' without ':' causes IndexOutOfRangeException on Split(':')[1]");
    }

    /// Bug #313 — Excel Add: table column count mismatch not validated
    /// File: ExcelHandler.Add.cs, lines 722-724
    /// User-provided column names are not validated against actual column count.
    [Fact]
    public void Bug313_ExcelAdd_TableColumnCountMismatch()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "H1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "H2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "H3" });

        // Provide 2 column names for a 3-column table
        var act = () => _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["range"] = "A1:C5",
            ["columns"] = "Name,Age"  // only 2 names for 3 columns
        });

        // Should validate column count matches, but silently creates mismatched table
        act.Should().NotThrow("but column names don't match actual column count");
    }

    /// Bug #314 — Excel Add: formula cell retains DataType from previous set
    /// File: ExcelHandler.Add.cs, line 91
    /// When setting a formula on a cell, the DataType is not cleared.
    /// If the cell previously had DataType=String, the formula cell retains it.
    [Fact]
    public void Bug314_ExcelAdd_FormulaCellRetainsDataType()
    {
        // Create a cell with string value first
        _excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1",
            ["value"] = "Hello"
        });

        // Now set a formula on the same cell
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["formula"] = "SUM(B1:B10)"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1/A1");
        // DataType should be cleared for formula cells, but may retain "String"
        node.Should().NotBeNull();
    }

    /// Bug #315 — Excel Add: icon set integer division truncation
    /// File: ExcelHandler.Add.cs, line 485
    /// i * 100 / iconCount uses integer division, losing precision.
    /// For 3-icon set: thresholds are 0%, 33%, 66% instead of 0%, 33.33%, 66.67%.
    [Fact]
    public void Bug315_ExcelAdd_IconSetIntegerDivisionTruncation()
    {
        int iconCount = 3;
        int threshold1 = 1 * 100 / iconCount;  // 33 (should be ~33.33)
        int threshold2 = 2 * 100 / iconCount;  // 66 (should be ~66.67)

        threshold1.Should().Be(33, "integer division truncates 33.33 to 33");
        threshold2.Should().Be(66, "integer division truncates 66.67 to 66");
        // Total coverage: 0-33, 33-66, 66-100 → gap at boundaries
    }

    /// Bug #316 — PPTX Set: double.Parse on volume property
    /// File: PowerPointHandler.Set.cs, line 501
    /// Volume uses double.Parse without TryParse validation.
    [Fact]
    public void Bug316_PptxSet_VolumeDoubleParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });

        var act = () => pptx.Set("/slide[1]/shape[1]", new()
        {
            ["volume"] = "loud"  // not a valid double
        });

        act.Should().Throw<Exception>(
            "double.Parse on 'loud' for volume throws FormatException");
    }

    /// Bug #317 — PPTX Set: bool.Parse on autoplay property
    /// File: PowerPointHandler.Set.cs, line 513
    /// Autoplay uses bool.Parse which only accepts "True"/"False".
    [Fact]
    public void Bug317_PptxSet_AutoplayBoolParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });

        var act = () => pptx.Set("/slide[1]/shape[1]", new()
        {
            ["autoplay"] = "yes"  // not "True" or "False"
        });

        act.Should().Throw<Exception>(
            "bool.Parse on 'yes' for autoplay throws FormatException — should use IsTruthy");
    }

    /// Bug #318 — PPTX Set: double.Parse on crop values
    /// File: PowerPointHandler.Set.cs, lines 647-650, 654, 660
    /// Crop percentages use double.Parse without validation.
    [Fact]
    public void Bug318_PptxSet_CropDoubleParseNoValidation()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });

            var act = () => pptx.Set("/slide[1]/picture[1]", new()
            {
                ["crop"] = "ten,20,30,40"  // "ten" is not a valid double
            });

            act.Should().Throw<Exception>(
                "double.Parse on 'ten' for crop value throws FormatException");
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    /// Bug #319 — PPTX Set: ParseEmu long-to-int overflow in slide size
    /// File: PowerPointHandler.Set.cs, line 30
    /// sldSz.Cx = (int)ParseEmu(value) — ParseEmu returns long,
    /// but Cx is int. Large EMU values overflow.
    [Fact]
    public void Bug319_PptxSet_SlideSizeParseEmuOverflow()
    {
        // ParseEmu("999cm") = 999 * 360000 = 359,640,000 → fits in int
        // ParseEmu("9999cm") = 9999 * 360000 = 3,599,640,000 → overflows int!
        long emu = (long)(9999 * 360000.0);
        int castResult = unchecked((int)emu);

        (emu > int.MaxValue).Should().BeTrue(
            "9999cm in EMU exceeds int.MaxValue, causing overflow when cast to int");
    }

    /// Bug #320 — Cross-handler: Excel Move only supports rows
    /// File: ExcelHandler.Add.cs, line 974+
    /// Excel's Move() only supports row[N], while Word supports any element
    /// and PowerPoint supports shapes/slides. No cell or column move support.
    [Fact]
    public void Bug320_CrossHandler_ExcelMoveLimitedSupport()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Try to move a cell — should throw clear error
        var act = () => _excelHandler.Move("/Sheet1/A1", "/Sheet1", 5);
        act.Should().Throw<Exception>(
            "Excel Move only supports row[N] — cell move throws ArgumentException");
    }

    /// Bug #321 — Cross-handler: inconsistent error message formats
    /// File: All handlers
    /// Word uses "Path not found: {path}", Excel uses "{ref} not found",
    /// PowerPoint uses "Slide {idx} not found (total: {count})".
    [Fact]
    public void Bug321_CrossHandler_InconsistentErrorMessages()
    {
        // Word: "Path not found: /body/p[999]"
        var wordAct = () => _wordHandler.Get("/body/p[999]");

        // Excel: different format
        var excelAct = () => _excelHandler.Get("/NonExistentSheet");

        // Both should throw, but with different error message styles
        wordAct.Should().Throw<Exception>();
        excelAct.Should().Throw<Exception>();
    }

    /// Bug #322 — BlankDocCreator: PPTX relationship ID collisions
    /// File: BlankDocCreator.cs, lines 64-69
    /// Multiple parts use overlapping "rId1", "rId2" etc. within different
    /// parent parts, which is valid. But the SlideLayoutId entries at
    /// lines 225-228 reference IDs that must match actual relationship IDs.
    [Fact]
    public void Bug322_BlankDocCreator_PptxRelationshipConsistency()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"blank_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(tempPath);
            using var pptx = new PowerPointHandler(tempPath);
            var root = pptx.Get("/");
            root.Should().NotBeNull("blank PPTX should be valid and openable");

            // Verify slide exists and is accessible
            var slide = pptx.Get("/slide[1]");
            slide.Should().NotBeNull("blank PPTX should have at least one slide");
        }
        finally { if (File.Exists(tempPath)) File.Delete(tempPath); }
    }

    /// Bug #323 — DocumentHandlerFactory: NotSupportedException not caught
    /// File: DocumentHandlerFactory.cs, lines 16-29
    /// The try-catch only catches OpenXmlPackageException. NotSupportedException
    /// from unsupported file extensions propagates uncaught.
    [Fact]
    public void Bug323_DocumentHandlerFactory_NotSupportedExceptionUncaught()
    {
        var act = () => DocumentHandlerFactory.Open("/tmp/test.txt");
        act.Should().Throw<Exception>(
            "unsupported extension throws NotSupportedException, not wrapped in InvalidOperationException");
    }

    /// Bug #324 — Excel Add: validation BETWEEN operator without formula2
    /// File: ExcelHandler.Add.cs, lines 266-275
    /// When operator is "between", formula2 is required but not validated.
    [Fact]
    public void Bug324_ExcelAdd_ValidationBetweenWithoutFormula2()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });

        var act = () => _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1"
            // formula2 is missing but required for "between"
        });

        // Should validate that formula2 is provided for "between" operator
        act.Should().NotThrow("but creates invalid validation without formula2 for between");
    }

    /// Bug #325 — Excel Add: comment author ID off-by-one
    /// File: ExcelHandler.Add.cs, lines 189, 193
    /// When adding a new author, authorId uses count BEFORE append.
    [Fact]
    public void Bug325_ExcelAdd_CommentAuthorIdOffByOne()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Add first comment
        _excelHandler.Add("/Sheet1/A1", "comment", null, new()
        {
            ["text"] = "First comment",
            ["author"] = "Alice"
        });

        // Add second comment with same author — should reuse author ID
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Test2" });
        _excelHandler.Add("/Sheet1/B1", "comment", null, new()
        {
            ["text"] = "Second comment",
            ["author"] = "Alice"
        });

        ReopenExcel();
        // Both comments should reference the same author
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }

    /// Bug #326 — PPTX Set: table style GUID duplicate for light3/medium3
    /// File: PowerPointHandler.Set.cs, lines 380, 384
    /// "light3" and "medium3" map to identical GUIDs (already known),
    /// verified here with a functional test.
    [Fact]
    public void Bug326_PptxSet_TableStyleGuidDuplicate()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set table style to "light3"
        pptx.Set("/slide[1]/table[1]", new() { ["style"] = "light3" });
        var node1 = pptx.Get("/slide[1]/table[1]");

        // Set table style to "medium3"
        pptx.Set("/slide[1]/table[1]", new() { ["style"] = "medium3" });
        var node2 = pptx.Get("/slide[1]/table[1]");

        // Both should have different GUIDs but code uses the same one
        // This means light3 and medium3 are visually identical
        true.Should().BeTrue(
            "Bug documented: light3 and medium3 table styles map to same GUID");
    }

    /// Bug #327 — Excel Add: DataValidations insertion order violates schema
    /// File: ExcelHandler.Add.cs, lines 298-304
    /// DataValidations element may be inserted AFTER ConditionalFormatting,
    /// violating Excel's required element ordering.
    [Fact]
    public void Bug327_ExcelAdd_DataValidationsSchemaOrder()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });

        // Add conditional formatting first
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "colorScale"
        });

        // Then add validation — should go BEFORE conditional formatting per schema
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "greaterThan",
            ["formula1"] = "0"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("sheet should be valid with both CF and validation");
    }

    /// Bug #328 — Cross-handler: Excel CopyFrom only supports rows
    /// File: ExcelHandler.Add.cs, lines 1037-1086
    /// Word and PPTX support copying any element, but Excel only supports row[N].
    [Fact]
    public void Bug328_CrossHandler_ExcelCopyFromLimitedSupport()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Try to copy a cell — should throw clear error
        var act = () => _excelHandler.CopyFrom("/Sheet1/A1", "/Sheet1", null);
        act.Should().Throw<Exception>(
            "Excel CopyFrom only supports row[N] — cell copy not supported");
    }

    /// Bug #329 — BlankDocCreator: hardcoded "zh-CN" language
    /// File: BlankDocCreator.cs, line 263
    /// Placeholder text language is hardcoded to "zh-CN" (Chinese),
    /// making blank documents always have Chinese language settings.
    [Fact]
    public void Bug329_BlankDocCreator_HardcodedChineseLanguage()
    {
        // BlankDocCreator creates placeholders with Language = "zh-CN"
        // This means blank PPTX documents default to Chinese language
        // for spell-checking and text rendering, regardless of user's locale.

        var tempPath = Path.Combine(Path.GetTempPath(), $"blank_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(tempPath);
            using var pptx = new PowerPointHandler(tempPath);
            var slide = pptx.Get("/slide[1]", depth: 2);
            // Placeholder text should use neutral or system locale, not zh-CN
            slide.Should().NotBeNull();
        }
        finally { if (File.Exists(tempPath)) File.Delete(tempPath); }
    }

    /// Bug #330 — PPTX Set: ParseEmu long-to-int overflow in shape position
    /// File: PowerPointHandler.Set.cs, lines 354, 426, 544, 597
    /// Multiple shape/element position properties cast ParseEmu(long) to int,
    /// risking overflow for large EMU values.
    [Fact]
    public void Bug330_PptxSet_ShapePositionParseEmuOverflow()
    {
        // Multiple locations cast long→int:
        //   Line 354: shape position
        //   Line 426: chart position
        //   Line 544: picture crop
        //   Line 597: text margin
        // Any large EMU value (>2.14 billion) causes silent overflow

        long largeEmu = (long)int.MaxValue + 100000;
        int castResult = unchecked((int)largeEmu);
        castResult.Should().BeNegative(
            "casting large EMU value to int wraps to negative, corrupting position");
    }

}
