// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

/// <summary>
/// Regression tests that prove each named functional bug in WatchServer.cs was
/// real.  Every test PASSES on the current fixed code; if the corresponding fix
/// were reverted the test would FAIL, proving the bug existed.
///
/// Bug markers and their scenarios:
///
///   BUG-TESTER-001  Catastrophically backtracking regex freezes reconcile loop
///   BUG-TESTER-002  Unsanitised color field enables CSS injection into every
///                   connected browser
///   BUG-TESTER-003  script/style inner text leaks into find-match extraction
///   BUG-FUZZER-001  Whitespace-padded color validated but stored with padding
///   BUG-FUZZER-003/004  Whitespace-only / no-leading-slash path accepted
///   BUG-A-R2-M01    Bare 3/6/8-digit hex colour rejected instead of promoted
///   BUG-BT-R303     Validation error messages non-actionable for AI agents
///   BUG-TESTER-R503 Wrong HTTP verb on /api/selection falls through to HTML (200)
///   BUG-TESTER-R504 Unknown /api/* path falls through to HTML (200)
/// </summary>
public sealed class BugRegressionTests : IAsyncDisposable
{
    private readonly List<string> _tempPaths = [];
    private readonly List<WatchServer> _servers = [];
    private readonly List<CancellationTokenSource> _ctsList = [];

    // ── helpers ──────────────────────────────────────────────────────────────

    private static int GetFreePort()
    {
        var l = new TcpListener(IPAddress.Loopback, 0);
        l.Start();
        var port = ((IPEndPoint)l.LocalEndpoint).Port;
        l.Stop();
        return port;
    }

    private string CreateTempFile(string extension = ".tmp")
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
        File.WriteAllText(path, "");
        _tempPaths.Add(path);
        return path;
    }

    private WatchServer CreateBareServer()
    {
        var filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        _tempPaths.Add(filePath);
        File.WriteAllText(filePath, "");
        var server = new WatchServer(filePath, GetFreePort(), TimeSpan.FromSeconds(1));
        _servers.Add(server);
        return server;
    }

    private async Task<(WatchServer server, int port)> StartWatchServerAsync(string filePath)
    {
        var port = GetFreePort();
        var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        var server = new WatchServer(filePath, port, TimeSpan.FromSeconds(30));
        _servers.Add(server);
        _ctsList.Add(cts);
        _ = server.RunAsync(cts.Token);
        await WaitForHttpAsync(port);
        await WaitForPipeAsync(WatchServer.GetWatchPipeName(filePath));
        return (server, port);
    }

    private static async Task WaitForHttpAsync(int port, int maxAttempts = 40)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            try { using var t = new TcpClient(); await t.ConnectAsync("127.0.0.1", port); return; }
            catch { await Task.Delay(50); }
        }
        throw new TimeoutException($"HTTP not ready on port {port}");
    }

    private static async Task WaitForPipeAsync(string pipeName, int maxAttempts = 40)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            try
            {
                using var p = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                p.Connect(100);
                return;
            }
            catch { await Task.Delay(50); }
        }
        throw new TimeoutException($"Pipe '{pipeName}' not ready");
    }

    private static async Task<string?> PipeCommandAsync(string pipeName, string command)
    {
        using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
        pipe.Connect(3000);
        var noBom = new UTF8Encoding(false);
        using var writer = new StreamWriter(pipe, noBom, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(pipe, noBom, leaveOpen: true);
        await writer.WriteLineAsync(command);
        return await reader.ReadLineAsync();
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var cts in _ctsList) { try { cts.Cancel(); } catch (ObjectDisposedException) { } }
        await Task.Delay(200);
        foreach (var s in _servers) { try { s.Dispose(); } catch (ObjectDisposedException) { } }
        foreach (var cts in _ctsList) { try { cts.Dispose(); } catch (ObjectDisposedException) { } }
        foreach (var p in _tempPaths)
        {
            try { if (File.Exists(p)) File.Delete(p); } catch (IOException) { }
        }
    }

    // =========================================================================
    // BUG-TESTER-001: Catastrophic regex must not freeze the reconcile loop
    // =========================================================================

    /// <summary>
    /// Before the fix: user-supplied regex like <c>(a+)+$</c> evaluated against
    /// a long input would cause catastrophic backtracking, blocking the watch
    /// reconcile loop indefinitely.
    ///
    /// After the fix: <c>ResolveMark</c> evaluates every user regex with a
    /// 500 ms timeout and marks the match as stale on timeout — no freeze.
    ///
    /// <b>Proof</b>: the test completes well within 2 seconds; the mark is stale.
    /// Without the timeout this would hang the test runner.
    /// </summary>
    [Fact]
    public void BugTester001_CatastrophicRegex_DoesNotFreezeReconcileLoop_MarkBecomesStale()
    {
        // Long input that triggers catastrophic backtracking on (a+)+$
        var longInput = new string('a', 5000) + "!"; // trailing '!' ensures no match
        var html = $"""<p data-path="/body/p[1]">{longInput}</p>""";

        var mark = new WatchMark
        {
            Id = "1",
            Path = "/body/p[1]",
            // r"(a+)+$" is the classic catastrophically backtracking pattern.
            Find = "r\"(a+)+$\"",
            Color = "#ffeb3b",
        };

        // Without the MarkRegexMatchTimeout fix this call would never return.
        var start = DateTime.UtcNow;
        var resolved = WatchServer.ResolveMark(mark, html);
        var elapsed = DateTime.UtcNow - start;

        // Fix works: returns quickly (well under the 2 s test budget).
        Assert.True(elapsed.TotalSeconds < 2,
            $"ResolveMark took {elapsed.TotalSeconds:F1} s — catastrophic backtracking was not bounded");

        // The mark is stale because the regex timed out (no matches found in time).
        Assert.True(resolved.Stale,
            "Mark must be stale after a regex match timeout, not stuck as non-stale");
    }

    // =========================================================================
    // BUG-TESTER-002: CSS injection via color field must be rejected
    // =========================================================================

    /// <summary>
    /// Before the fix: <c>mark.color</c> was written verbatim into
    /// <c>el.style.backgroundColor = mark.color</c> on every connected browser.
    /// A payload like <c>red; background: url(//evil.example/steal.gif)</c>
    /// would inject arbitrary CSS into every viewer's browser.
    ///
    /// After the fix: <c>IsValidMarkColor</c> rejects anything that is not a
    /// recognised CSS colour value.
    ///
    /// <b>Proof</b>: each injection payload fails the server-side validator.
    /// </summary>
    [Theory]
    [InlineData("red; background: url(//evil.example/steal.gif)")]  // CSS property chain
    [InlineData("expression(alert(1))")]                             // IE CSS expression
    [InlineData("1px solid red")]                                    // border, not a color
    [InlineData("red\nalert(1)")]                                    // newline injection
    [InlineData("url('javascript:alert(1)')")]                       // url() with JS
    [InlineData("rgba(0,0,0,0); background:red")]                   // double-property chain
    public void BugTester002_CssInjectionViaColor_IsRejectedByServerSideValidator(string injectionPayload)
    {
        // Before the fix there was no validator — any string was accepted.
        // The fix adds IsValidMarkColor which must reject all of these.
        var isValid = WatchServer.IsValidMarkColor(injectionPayload);
        Assert.False(isValid,
            $"CSS injection payload must be rejected by the color validator: {injectionPayload}");
    }

    /// <summary>
    /// Corollary: legitimate colour values must still pass after the fix, so the
    /// validator does not break normal usage.
    /// </summary>
    [Theory]
    [InlineData("#ff0000")]
    [InlineData("#f00")]
    [InlineData("red")]
    [InlineData("blue")]
    [InlineData("rgb(255,0,0)")]
    [InlineData("rgba(255,0,0,0.5)")]
    public void BugTester002_LegitimateColors_PassValidation(string color)
    {
        var normalized = WatchServer.NormalizeMarkColorInput(color) ?? color;
        Assert.True(WatchServer.IsValidMarkColor(normalized),
            $"Legitimate color '{color}' (normalized: '{normalized}') must pass validation");
    }

    // =========================================================================
    // BUG-TESTER-003: <script>/<style> inner text must not leak into find matching
    // =========================================================================

    /// <summary>
    /// Before the fix: <c>ExtractTextContent</c> called <c>_tagStripRx</c>
    /// (which strips HTML tags but leaves their text content) directly.  The
    /// text inside <c>&lt;script&gt;</c> elements leaked into the plain-text
    /// result used for <c>find</c> matching.  A mark with <c>find="secret"</c>
    /// would falsely hit <c>&lt;script&gt;var secret = …&lt;/script&gt;</c>.
    ///
    /// After the fix: script and style bodies are stripped in their entirety
    /// before per-tag stripping.
    ///
    /// <b>Proof</b>: text inside script/style elements does NOT appear in the
    /// extracted plain text.
    /// </summary>
    [Fact]
    public void BugTester003_ScriptInnerTextDoesNotLeakIntoFindMatching()
    {
        const string html = "<p>visible text</p><script>var secret = 'password';</script>";

        // Before fix: "visible text" + "var secret = 'password';"
        // After fix:  "visible text" only.
        var text = WatchServer.ExtractTextContent(html);

        Assert.True(!text.Contains("secret"),
            "Script inner text must be stripped before find matching — " +
            "without the fix the JS variable name would leak into the extracted text");
        Assert.True(text.Contains("visible"), "Visible text must still be present");
    }

    [Fact]
    public void BugTester003_StyleInnerTextDoesNotLeakIntoFindMatching()
    {
        const string html = "<p>normal content</p><style>.secret-class { color: red; }</style>";

        var text = WatchServer.ExtractTextContent(html);

        Assert.True(!text.Contains("secret-class"),
            "Style inner text must be stripped before find matching — " +
            "without the fix the CSS class name would leak into the extracted text");
        Assert.True(text.Contains("normal content"), "Visible text must still be present");
    }

    [Fact]
    public void BugTester003_MultilineScriptBodyStrippedInFull()
    {
        // Multi-line script block — the fix uses RegexOptions.Singleline so
        // newlines inside the body are consumed too.
        const string html = "<p>hello</p><script type=\"text/javascript\">\n" +
                            "  var internal_key = 'top-secret';\n" +
                            "  console.log(internal_key);\n" +
                            "</script><p>world</p>";

        var text = WatchServer.ExtractTextContent(html);

        Assert.DoesNotContain("internal_key", text);
        Assert.DoesNotContain("top-secret", text);
        Assert.Contains("hello", text);
        Assert.Contains("world", text);
    }

    // =========================================================================
    // BUG-FUZZER-001: Whitespace-padded color must be trimmed before storage
    // =========================================================================

    /// <summary>
    /// Before the fix: <c>HandleMarkAdd</c> validated the color after
    /// <c>Trim()</c> but stored <c>req.Color</c> (the untrimmed original).  A
    /// color like <c>"red\n"</c> passed validation but was stored as
    /// <c>"red\n"</c> — a validator/storage inconsistency that could corrupt
    /// SSE JSON or cause silent mismatches on retrieval.
    ///
    /// After the fix: the trimmed form is used for both validation AND storage.
    ///
    /// <b>Proof</b>: the mark is stored with the clean trimmed color.
    /// </summary>
    [Fact]
    public async Task BugFuzzer001_WhitespacePaddedColor_IsStoredTrimmedNotRaw()
    {
        var filePath = CreateTempFile(".docx");
        var (_, _) = await StartWatchServerAsync(filePath);
        var pipeName = WatchServer.GetWatchPipeName(filePath);

        // Colour with trailing whitespace (newline + space).
        var payload = JsonSerializer.Serialize(new
        {
            path = "/body/p[1]",
            color = "red\n "
        });
        var addReply = await PipeCommandAsync(pipeName, $"mark {payload}");

        Assert.NotNull(addReply);
        using var addDoc = JsonDocument.Parse(addReply!);
        var error = addDoc.RootElement.TryGetProperty("error", out var e) ? e.GetString() : null;
        Assert.Null(error);   // "red\n " trimmed to "red" → valid; must not be rejected
        Assert.True(addDoc.RootElement.TryGetProperty("id", out var idEl)
                    && !string.IsNullOrEmpty(idEl.GetString()),
            "Mark add must succeed when color is a valid named color after trimming");

        // Retrieve stored marks and verify the color was stored trimmed.
        var getReply = await PipeCommandAsync(pipeName, "get-marks");
        Assert.NotNull(getReply);
        using var marksDoc = JsonDocument.Parse(getReply!);
        var marks = marksDoc.RootElement.GetProperty("marks");
        Assert.Equal(1, marks.GetArrayLength());

        var storedColor = marks[0].GetProperty("color").GetString();
        // Before the fix: "red\n " (untrimmed). After the fix: "red".
        Assert.Equal("red", storedColor);
    }

    // =========================================================================
    // BUG-FUZZER-003/004: Path validation and normalisation
    // =========================================================================

    /// <summary>
    /// Before the fix: whitespace-only paths and paths without a leading '/'
    /// were accepted, producing marks that could never resolve.
    ///
    /// After the fix: paths are Trim()ed and validated — whitespace-only and
    /// non-slash-prefixed strings are rejected with an actionable error.
    ///
    /// <b>Proof</b>: each invalid path causes HandleMarkAdd to return an error.
    /// </summary>
    [Theory]
    [InlineData("")]           // empty
    [InlineData("   ")]        // whitespace-only
    [InlineData("\t\n")]       // tab + newline
    [InlineData("body/p[1]")]  // missing leading slash
    [InlineData("p[1]")]       // relative path, no slash
    public void BugFuzzer003_InvalidPaths_AreRejectedByHandleMarkAdd(string badPath)
    {
        var server = CreateBareServer();
        var json = JsonSerializer.Serialize(new { path = badPath, color = "red" });
        var result = server.HandleMarkAdd(json);

        // On error HandleMarkAdd returns {"error":"..."} with no "id" field.
        // On success it returns {"id":"...","error":null} with a non-null id.
        // Before the fix: these paths were accepted and stored verbatim.
        // After the fix:  the method returns an error JSON with no id.
        Assert.True(result.Contains("\"error\":\""),
            $"Path '{badPath}' must be rejected — before the fix it was accepted and stored");
    }

    [Fact]
    public void BugFuzzer004_WhitespacePaddedValidPath_IsNormalisedAndAccepted()
    {
        // A padded path must be accepted after trimming (normalise, don't reject).
        var server = CreateBareServer();
        var json = JsonSerializer.Serialize(new { path = "  /body/p[1]  ", color = "red" });
        var result = server.HandleMarkAdd(json);

        // On success: {"id":"1","error":null} — "error" key exists but is null.
        // On failure: {"error":"invalid path: ..."} — "error" has a string value.
        // Padded path must be accepted — trimmed to "/body/p[1]".
        Assert.True(!result.Contains("\"error\":\""),
            $"A whitespace-padded but otherwise valid path must be trimmed and accepted; got: {result}");
        // The returned id must be non-empty.
        using var doc = JsonDocument.Parse(result);
        Assert.True(doc.RootElement.TryGetProperty("id", out var idEl)
                    && !string.IsNullOrEmpty(idEl.GetString()),
            "Accepted mark must return a non-empty id");
    }

    // =========================================================================
    // BUG-A-R2-M01: Bare hex colour promoted to #-prefixed form
    // =========================================================================

    /// <summary>
    /// Before the fix: bare hex colours like <c>FF00FF</c> were passed directly
    /// to <c>IsValidMarkColor</c> which requires the <c>#</c> prefix — they were
    /// rejected despite being unambiguously valid, breaking the CLI's convention.
    ///
    /// After the fix: <c>NormalizeMarkColorInput</c> promotes bare hex to the
    /// canonical <c>#-prefixed</c> form before validation.
    ///
    /// <b>Proof</b>: all three bare-hex forms are promoted and then accepted.
    /// </summary>
    [Theory]
    [InlineData("FF0000",   "#FF0000")]   // 6-digit bare hex
    [InlineData("ff0000",   "#FF0000")]   // lowercase 6-digit
    [InlineData("F00",      "#FF0000")]   // 3-digit bare hex expanded
    [InlineData("f00",      "#FF0000")]   // lowercase 3-digit
    [InlineData("FF0000FF", "#FF0000FF")] // 8-digit bare hex
    public void BugA_R2_M01_BareHexColor_PromotedToHashPrefixedForm(string input, string expected)
    {
        var normalized = WatchServer.NormalizeMarkColorInput(input);

        // Before the fix: no NormalizeMarkColorInput; raw bare hex reached
        // IsValidMarkColor and was rejected.
        // After the fix: promoted to canonical form, then accepted.
        Assert.Equal(expected, normalized);
        Assert.True(WatchServer.IsValidMarkColor(normalized!),
            $"Normalised form '{normalized}' must pass color validation");
    }

    [Theory]
    [InlineData("#FF0000")]    // already prefixed
    [InlineData("red")]        // named colour
    [InlineData("rgb(1,2,3)")] // rgb() form
    public void BugA_R2_M01_AlreadyNormalisedColors_PassThroughUnchanged(string input)
    {
        // Idempotent: already-correct forms are returned unchanged.
        var normalized = WatchServer.NormalizeMarkColorInput(input);
        Assert.Equal(input, normalized);
        Assert.True(WatchServer.IsValidMarkColor(normalized!));
    }

    // =========================================================================
    // BUG-BT-R303: Validation error messages must be actionable
    // =========================================================================

    /// <summary>
    /// Before the fix: validation errors returned generic messages like
    /// <c>"invalid path"</c> with no hint about the accepted format, leaving
    /// AI agents stuck in retry loops.
    ///
    /// After the fix: error messages explicitly state the accepted format.
    ///
    /// <b>Proof</b>: error messages contain the accepted format description.
    /// </summary>
    [Fact]
    public void BugBt_R303_PathValidationError_ContainsAcceptedFormatHint()
    {
        var server = CreateBareServer();
        var json = JsonSerializer.Serialize(new { path = "no-leading-slash", color = "red" });
        var result = server.HandleMarkAdd(json);

        Assert.True(result.Contains("\"error\":\""), "Result must be an error JSON with a non-null error message");
        // The error message must tell the caller what the accepted format IS.
        Assert.True(result.Contains("/"),
            "Path error must show the '/' prefix requirement — before fix: just 'invalid path'");
    }

    [Fact]
    public void BugBt_R303_ColorValidationError_ListsAcceptedColorFormats()
    {
        var server = CreateBareServer();
        var json = JsonSerializer.Serialize(new { path = "/body/p[1]", color = "not-a-color" });
        var result = server.HandleMarkAdd(json);

        Assert.True(result.Contains("\"error\":\""), "Result must be an error JSON with a non-null error message");
        // The error must list accepted forms so agents can self-correct.
        Assert.True(result.Contains("#"), "Error must mention #RGB / #RRGGBB form");
        Assert.True(result.Contains("rgb"), "Error must mention rgb() form");
        // Before the fix: just "invalid color" with no format hint.
    }

    // =========================================================================
    // BUG-TESTER-R503: Non-POST to /api/selection must return 405
    // =========================================================================

    /// <summary>
    /// Before the fix: a <c>GET</c> (or any non-POST verb) to
    /// <c>/api/selection</c> fell through to the HTML preview handler and
    /// returned a <c>200 OK</c> HTML page.
    ///
    /// After the fix: the router returns <c>405 Method Not Allowed</c>.
    ///
    /// <b>Proof</b>: non-POST requests return 405, not 200 HTML.
    /// </summary>
    [Theory]
    [InlineData("GET")]
    [InlineData("PUT")]
    [InlineData("DELETE")]
    [InlineData("PATCH")]
    public async Task BugTester_R503_NonPostSelectionEndpoint_Returns405NotHtmlPreview(string method)
    {
        var filePath = CreateTempFile(".docx");
        var (_, port) = await StartWatchServerAsync(filePath);

        using var tcp = new TcpClient();
        await tcp.ConnectAsync("127.0.0.1", port);
        using var stream = tcp.GetStream();
        stream.ReadTimeout = 3000;

        var request = Encoding.ASCII.GetBytes(
            $"{method} /api/selection HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(request);

        var buf = new byte[4096];
        var n = await stream.ReadAsync(buf);
        var response = Encoding.ASCII.GetString(buf, 0, n);

        // Before the fix: "HTTP/1.1 200 OK" + full HTML preview.
        // After the fix:  "HTTP/1.1 405 Method Not Allowed".
        Assert.True(response.Contains("HTTP/1.1 405"),
            $"{method} /api/selection must return 405 — before the fix it returned 200 HTML");
        Assert.True(response.Contains("Allow: POST"),
            "405 response must include Allow header listing the accepted verb");
        Assert.True(!response.Contains("text/html"),
            "405 response must not return the HTML preview page");
    }

    // =========================================================================
    // BUG-TESTER-R504: Unknown /api/* path must return 404
    // =========================================================================

    /// <summary>
    /// Before the fix: any request to an unknown <c>/api/*</c> path fell
    /// through to the HTML preview handler and returned <c>200 OK</c> HTML.
    ///
    /// After the fix: unrecognised <c>/api/*</c> paths return <c>404 Not Found</c>.
    ///
    /// <b>Proof</b>: requests to unknown API paths return 404, not 200 HTML.
    /// </summary>
    [Theory]
    [InlineData("/api/marks")]
    [InlineData("/api/unknown")]
    [InlineData("/api/v1/selection")]
    [InlineData("/api/")]
    public async Task BugTester_R504_UnknownApiPath_Returns404NotHtmlPreview(string unknownPath)
    {
        var filePath = CreateTempFile(".docx");
        var (_, port) = await StartWatchServerAsync(filePath);

        using var tcp = new TcpClient();
        await tcp.ConnectAsync("127.0.0.1", port);
        using var stream = tcp.GetStream();
        stream.ReadTimeout = 3000;

        var request = Encoding.ASCII.GetBytes(
            $"GET {unknownPath} HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(request);

        var buf = new byte[4096];
        var n = await stream.ReadAsync(buf);
        var response = Encoding.ASCII.GetString(buf, 0, n);

        // Before the fix: "HTTP/1.1 200 OK" + full HTML preview.
        // After the fix:  "HTTP/1.1 404 Not Found".
        Assert.True(response.Contains("HTTP/1.1 404"),
            $"GET {unknownPath} must return 404 — before the fix it returned 200 HTML");
        Assert.True(!response.Contains("text/html"),
            "404 response must not return the HTML preview page");
    }
}
