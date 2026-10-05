using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Newtonsoft.Json.Linq;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;


namespace TestAutomationFramework.Framework
{
    public class TestSuiteRunner
    {
        private IWebDriver? _driver;
        private WebDriverWait? _wait;
        private string? _lastFrameCss;
        private string? _dataFilesDir;
        private bool _windowsAuth;
        private string? _authServerAllowlist;


        public class TestResult
        {
            public string FileName { get; set; } = "";
            public bool Passed { get; set; }
            public List<string> Logs { get; set; } = new();
            public List<string> Errors { get; set; } = new();
            public string OutputFilePath { get; set; } = "";
            public string ScreenshotPath { get; set; } = "";
        }

        private class TestStep
        {
            public int Row { get; set; }
            public string Command { get; set; } = "";
            public string? Target { get; set; }
            public string? Value { get; set; }
            public string? Description { get; set; }
            public List<string[]>? Selectors { get; set; } // For DevTools JSON steps and CSV of selectors from Excel
        }

        public TestResult Run(string excelPath, bool headless = false, bool incognito = false, int stepDelayMs = 0,
            bool windowsAuth = false, string? authServerAllowlist = null)
        {
            _authServerAllowlist = BrowserOptionsFactory.ValidateAuthentication(windowsAuth, incognito, authServerAllowlist);
            BrowserOptionsFactory.EnsureWindowsAuthenticationSupported(windowsAuth);
            _windowsAuth = windowsAuth;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var result = new TestResult { FileName = Path.GetFileName(excelPath) };

            var testDir = Path.GetDirectoryName(excelPath)
                ?? throw new InvalidOperationException("Test file directory could not be resolved.");
            
            var inputRoot = FindAncestorNamed(testDir, "Input")
                ?? throw new InvalidOperationException($"Test file must be a child of the Input folder. File: {excelPath}");
            
            var projectRoot = Directory.GetParent(inputRoot)?.FullName
                ?? throw new InvalidOperationException($"Input folder has no parent. Input: {inputRoot}");
            
            var relativeDir = Path.GetRelativePath(inputRoot, testDir);
            var outputDir = relativeDir == "."
                ? Path.Combine(projectRoot, "Output")
                : Path.Combine(projectRoot, "Output", relativeDir);
            
            Directory.CreateDirectory(outputDir);
            
            _dataFilesDir = Path.Combine(projectRoot, "DataFiles");
            Directory.CreateDirectory(_dataFilesDir);

            string baseName = Path.GetFileNameWithoutExtension(excelPath);
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
            var tokens = new TokenResolver(
                emailPrefix: "vsc.qa.",
                emailDomain: "@example.com",
                usernamePrefix: "vsc_user_"
               );


            using var inputStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var inputPackage = new ExcelPackage(inputStream);
            ExcelWorksheet inputSheet;
            try
            {
                if (inputPackage.Workbook.Worksheets.Count == 0)
                    throw new InvalidDataException($"No worksheet found in test file: {excelPath}");

                inputSheet = inputPackage.Workbook.Worksheets[0];
            }
            catch (IOException ex)
            {
                throw new IOException($"Failed to read test file '{excelPath}'. Save/close Excel and retry.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Invalid or unreadable Excel test file '{excelPath}'.", ex);
            }

            using var resultPackage = new ExcelPackage();
            var sheet = resultPackage.Workbook.Worksheets.Add("Sheet1", inputSheet);

            // Build steps from the copied sheet
            var steps = new List<TestStep>();
            int row = 2;
            while (!string.IsNullOrWhiteSpace(sheet.Cells[row, 3].Text))
            {
                // NOTE: if your locator text can contain commas, change Split(',') to Split(';')
                var rawTarget = sheet.Cells[row, 4].Text?.Trim() ?? "";
                var selectorList = rawTarget.Length == 0
     ? null
                 : new List<string[]> { rawTarget.Split(',').Select(s => s.Trim()).ToArray() };


                steps.Add(new TestStep
                {
                    Row = row,
                    Command = sheet.Cells[row, 3].Text.Trim(),
                    Target = rawTarget,
                    Selectors = selectorList,
                    Value = sheet.Cells[row, 5].Text.Trim(),
                    Description = sheet.Cells[row, 2].Text.Trim()
                });

                row++;
            }

            StartDriver(headless, incognito);

            bool anyRequiredFailure = false;
            bool anyOptionalFailure = false;

            foreach (var step in steps)
            {
                //1) Optional command parsing (leading '*' only)
                var (cleanCmd, isOptional) = ParseOptionalCommand(step.Command);

                //2) Replace tokens in Value/Target (common use-case: registration data)
                step.Value = tokens.ReplaceTokens(step.Value);
                step.Target = tokens.ReplaceTokens(step.Target);

                //3) Replace tokens in selectors too (safe; usually not needed but harmless)
                if (step.Selectors != null)
                {
                    for (int i = 0; i < step.Selectors.Count; i++)
                    {
                        for (int j = 0; j < step.Selectors[i].Length; j++)
                        {
                            step.Selectors[i][j] = tokens.ReplaceTokens(step.Selectors[i][j]);
                        }
                    }
                }

                //4) Override command to clean command so Execute() switch matches
                step.Command = cleanCmd;

                try
                {
                    Execute(step);
                    sheet.Cells[step.Row, 1].Value = "Passed";
                    result.Logs.Add($"Step {step.Row - 1} passed.");
                    if (headless && stepDelayMs > 0)
                    {
                        System.Threading.Thread.Sleep(stepDelayMs);
                    }
                }
                catch (Exception ex)
                {
                    //OPTIONAL failure: do not stop test
                    if (isOptional)
                    {
                        anyOptionalFailure = true;

                        string warnMsg = !string.IsNullOrEmpty(step.Description)
                            ? $"Optional step {step.Row - 1} failed: {step.Description} — {ex.Message}"
                            : $"Optional step {step.Row - 1} failed: {step.Command} → {step.Target} — {ex.Message}";

                        sheet.Cells[step.Row, 1].Value = "Optional Failed";
                        result.Errors.Add(warnMsg);

                        //(Optional) take screenshot for optional failures too


                        continue;
                    }

                    //REQUIRED failure: existing behavior (screenshot + stop)
                    anyRequiredFailure = true;

                    //Save screenshot
                    string screenshotDir = Path.Combine(outputDir, "TestFailures");
                    Directory.CreateDirectory(screenshotDir);
                    string screenshotName = $"{baseName}_Step{step.Row - 1}_Failure_{timestamp}.png";
                    string screenshotPath = Path.Combine(screenshotDir, screenshotName);

                    try
                    {
                        Screenshot ss = ((ITakesScreenshot)_driver!).GetScreenshot();
                        ss.SaveAsFile(screenshotPath);
                        result.ScreenshotPath = screenshotPath;
                    }
                    catch (Exception ssex)
                    {
                        result.Errors.Add($"Failed to save screenshot: {ssex.Message}");
                    }

                    string errorMsg = !string.IsNullOrEmpty(step.Description)
                        ? $"Step {step.Row - 1} failed: {step.Description} — {ex.Message}"
                        : $"Step {step.Row - 1} failed: {step.Command} → {step.Target} — {ex.Message}";

                    sheet.Cells[step.Row, 1].Value = "Failed";
                    result.Errors.Add(errorMsg);
                    result.Passed = false;
                    break;
                }
            }

            //Write summary and save results workbook
            int resultRow = row + 2;
            sheet.Cells[resultRow, 1].Value = "Result:";
            sheet.Cells[resultRow, 2].Value =
    anyRequiredFailure ? "Failed" :
    anyOptionalFailure ? "PASSED_WITH_WARNINGS" :
    "Passed";

            string outputName = $"{baseName}_{timestamp}.xlsx";
            string outputPath = Path.Combine(outputDir, outputName);
            result.OutputFilePath = outputPath;
            resultPackage.SaveAs(new FileInfo(outputPath));

            _driver?.Quit();
            _driver?.Dispose();

            result.Passed = !anyRequiredFailure;
            return result;
        }
        private static string? FindAncestorNamed(string startPath, string folderName)
        {
            var dir = new DirectoryInfo(startPath);
            while (dir != null)
            {
                if (string.Equals(dir.Name, folderName, StringComparison.OrdinalIgnoreCase))
                    return dir.FullName;

                dir = dir.Parent;
            }

            return null;
        }
        private void StartDriver(bool headless, bool incognito)
        {
            // Allow Selenium/Selenium Manager behavior when explicit path is not used.
            Environment.SetEnvironmentVariable("SE_DISABLE_DRIVER_MANAGEMENT", null);

            //The exe folder (where your app EXE and chromedriver.exe live)
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string driverExe = Path.Combine(exeDir, "chromedriver.exe");

            var options = BrowserOptionsFactory.Create(headless, incognito, _windowsAuth, _authServerAllowlist);

            try
            {
                if (File.Exists(driverExe))
                {
                    var service = ChromeDriverService.CreateDefaultService(exeDir, "chromedriver.exe");
                    service.HideCommandPromptWindow = true;
                    _driver = new ChromeDriver(service, options, TimeSpan.FromSeconds(60));
                }
                else
                {
                    _driver = new ChromeDriver(options);
                }
            }
            catch (WebDriverException ex) when (
                ex.Message.Contains("only supports Chrome version", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("session not created", StringComparison.OrdinalIgnoreCase))
            {
                _driver?.Quit();
                _driver?.Dispose();

                // Resolve/download matching driver for installed Chrome and retry explicitly.
                string managedDriverPath = new DriverManager().SetUpDriver(
                    new ChromeConfig(),
                    VersionResolveStrategy.MatchingBrowser);

                var managedService = ChromeDriverService.CreateDefaultService(
                    Path.GetDirectoryName(managedDriverPath)!,
                    Path.GetFileName(managedDriverPath));
                managedService.HideCommandPromptWindow = true;

                _driver = new ChromeDriver(managedService, options, TimeSpan.FromSeconds(60));
            }

            //Extra maximize call for reliability in normal mode
            if (!headless)
                _driver.Manage().Window.Maximize();

            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
        }

        private void Execute(TestStep s)
        {
            string command = s.Command?.Trim().ToLowerInvariant() ?? "";

            switch (command)
            {
                case "open":
                    _driver!.Navigate().GoToUrl(s.Target!);
                    // optional helpers you already have:
                    //pophandler disabled (no longer required)
                    //PopupHandler.HandleCRMWelcomePopup(_driver!);
                    if (IsCrmContext())
                    {
                        TryEnterCrmMainFrame();
                    }

                    break;

                case "click":
                    {
                        var target = (s.Selectors != null ? Locate(s.Selectors) : Locate(new List<string[]> { new[] { s.Target! } }));

                        if (ShouldUseSalesforceMode(s.Selectors ?? new List<string[]> { new[] { s.Target ?? "" } }))
                        {
                            LightningSafeClick(target);
                        }
                        else
                        {
                            target.Click();
                        }

                        break;
                    }


                case "doubleclick":
                    var action = new OpenQA.Selenium.Interactions.Actions(_driver!); // null-forgiving
                    var dcTarget = s.Selectors != null ? Locate(s.Selectors) : Locate(new List<string[]> { new[] { s.Target! } });
                    action.DoubleClick(dcTarget).Perform();
                    break;

                case "type":
                case "change":
                    {
                        var elem = s.Selectors != null
                            ? Locate(s.Selectors)
                            : Locate(new List<string[]> { new[] { s.Target! } });

                        // Detect DataFiles token OR DevTools fakepath upload
                        string? uploadPath = TryResolveUploadPath(s.Value);

                        if (uploadPath != null)
                        {
                            // File upload: do NOT Clear()
                            elem.SendKeys(uploadPath);
                            break;
                        }

                        // Normal typing
                        try
                        {
                            elem.Clear();
                        }
                        catch
                        {
                            // Some Lightning inputs/textarea wrappers can be weird with Clear()
                            // If Clear() fails, fall back to CTRL+A + Backspace
                            elem.SendKeys(Keys.Control + "a");
                            elem.SendKeys(Keys.Backspace);
                        }

                        elem.SendKeys(s.Value ?? "");

                        // Salesforce-only: dispatch input/change to ensure Lightning validation + persistence
                        if (ShouldUseSalesforceMode(s.Selectors ?? new List<string[]> { new[] { s.Target ?? "" } }))
                        {
                            DispatchLightningInputAndChange(elem);
                            // NOTE: we are NOT auto-tabbing here (keeps CRM + general stability)
                            // If a specific field still needs commit, use a sendKeys step with ${KEY_TAB} in Excel.
                        }

                        break;
                    }


                case "sendkeys":
                    var el = s.Selectors != null ? Locate(s.Selectors) : Locate(new List<string[]> { new[] { s.Target! } });
                    var keys = (s.Value ?? "")
                        .Replace("${KEY_ENTER}", Keys.Enter)
                        .Replace("${KEY_TAB}", Keys.Tab)
                        .Replace("${KEY_ESC}", Keys.Escape);
                    el.SendKeys(keys);
                    break;

                case "sleep":
                case "wait":
                    System.Threading.Thread.Sleep(ParseSleepDurationMs(s));
                    break;

                case "waitforelement":
                    WaitForElementAssertion(s);
                    break;


                case "selectframe":
                    if (s.Target!.StartsWith("index="))
                        _driver!.SwitchTo().Frame(int.Parse(s.Target["index=".Length..]));
                    else if (s.Target == "relative=parent")
                        _driver!.SwitchTo().ParentFrame();
                    else
                        _driver!.SwitchTo().Frame(Locate(new List<string[]> { new[] { s.Target! } }));
                    break;

                case "close":
                    _driver!.Close();
                    break;

                case "select":
                    {
                        var selectors = s.Selectors ?? new List<string[]> { new[] { s.Target! } };

                        // 1) Determine the intended value (support legacy tests where value is stuffed into Target as text=...)
                        string val = (s.Value ?? "").Trim();
                        if (string.IsNullOrWhiteSpace(val))
                        {
                            string? fromTarget = ExtractSelectValueFromTarget(selectors);
                            if (!string.IsNullOrWhiteSpace(fromTarget))
                                val = fromTarget;
                        }

                        // 2) For non-Salesforce, prefer locating an actual <select> using css/id/xpath selectors first
                        IWebElement dropdown;

                        bool isSf = ShouldUseSalesforceMode(selectors);

                        if (!isSf)
                        {
                            dropdown = LocatePreferFormControl(selectors);
                        }
                        else
                        {
                            dropdown = Locate(selectors);
                        }

                        // 3) If it's a real <select>, use SelectElement (existing behavior)
                        if (string.Equals(dropdown.TagName, "select", StringComparison.OrdinalIgnoreCase))
                        {
                            var select = new SelectElement(dropdown);

                            if (val.StartsWith("label=", StringComparison.OrdinalIgnoreCase))
                                select.SelectByText(val["label=".Length..].Trim());
                            else if (val.StartsWith("value=", StringComparison.OrdinalIgnoreCase))
                                select.SelectByValue(val["value=".Length..].Trim());
                            else if (val.StartsWith("index=", StringComparison.OrdinalIgnoreCase) &&
                                     int.TryParse(val["index=".Length..], out int i))
                                select.SelectByIndex(i);
                            else
                            {
                                // legacy: if value is numeric, treat as value=
                                if (val.All(char.IsDigit))
                                    select.SelectByValue(val);
                                else
                                    select.SelectByText(val);
                            }

                            break;
                        }

                        // 4) Salesforce Lightning combobox selection
                        if (isSf)
                        {
                            string desiredText = NormalizeSelectValueToText(val);
                            SelectLightningComboboxOption(dropdown, desiredText);
                            break;
                        }

                        // 5) Generic fallback for non-<select> dropdowns (don’t hard-fail old tests)
                        // Click then choose option by visible text (exact match)
                        dropdown.Click();
                        string optionText = NormalizeSelectValueToText(val);
                        string optionXpath = $"//div[@role='listbox']//*[normalize-space(.)={QuoteForXpath(optionText)}]";
                        Locate(new List<string[]> { new[] { "xpath=" + optionXpath } }).Click();
                        break;
                    }



                case "hover":
                    var hoverTarget = s.Selectors != null ? Locate(s.Selectors) : Locate(new List<string[]> { new[] { s.Target! } });
                    new OpenQA.Selenium.Interactions.Actions(_driver!).MoveToElement(hoverTarget).Perform();
                    break;

                case "scroll":
                    if (!string.IsNullOrEmpty(s.Value) && s.Value.Contains(","))
                    {
                        var parts = s.Value.Split(',');
                        if (int.TryParse(parts[0], out int x) && int.TryParse(parts[1], out int y))
                            ((IJavaScriptExecutor)_driver!).ExecuteScript($"window.scrollTo({x}, {y});");
                    }
                    break;

                case "waitforexpression":
                    string expression = s.Value ?? "false";
                    _wait!.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript($"return ({expression});") is true);
                    break;

                case "assertpresent":
                    _ = Locate(s.Selectors ?? new List<string[]> { new[] { s.Target! } });
                    break;

                case "assertnotpresent":
                    try
                    {
                        _ = Locate(s.Selectors ?? new List<string[]> { new[] { s.Target! } });
                        throw new Exception($"Expected no element to be present for target '{s.Target}', but an element was found.");
                    }
                    catch (NoSuchElementException)
                    {
                        break;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        break;
                    }
                    break;

                case "asserttext":
                    var textElem = Locate(s.Selectors ?? new List<string[]> { new[] { s.Target! } });
                    var actualText = textElem.Text;
                    if (!actualText.Contains(s.Value ?? "", StringComparison.OrdinalIgnoreCase))
                        throw new Exception($"Expected text '{s.Value}'. but found '{actualText}'");
                    break;

                case "asserturl":
                    string currentUrl = _driver!.Url;
                    if (!string.Equals(currentUrl, s.Value, StringComparison.OrdinalIgnoreCase))
                        throw new Exception($"Expected URL '{s.Value}' but found '{currentUrl}'");
                    break;

                case "switchwindow":
                    List<String> windowtitles = new List<String> { };
                    bool found = false;
                    String originalwindow = _driver!.CurrentWindowHandle;
                    string targetTitle = s.Value ?? "";
                            var windows = _driver!.WindowHandles;
                            foreach (var handle in windows)
                            {
                                _driver.SwitchTo().Window(handle);
                                windowtitles.Add(_driver.Title);
                        if (_driver.Title == targetTitle)
                        {
                            found = true;
                            break;
                        }                     
                            }
                    if (found == false)
                        if (found == false)
                           throw new Exception($"No window found with title containing '{targetTitle}', List of all window titles, seperated by commas: '{string.Join(", ", windowtitles)}'");


                    break;


                default:
                    throw new NotSupportedException($"Unknown command: '{s.Command}'");
            }
        }
        // ===== DevTools waitForElement assertion support =====

        private sealed class WaitForElementCfg
        {
            public int TimeoutMs { get; set; } = 10000;          // default 10s
            public bool? Visible { get; set; }                   // true/false if specified
            public int? Count { get; set; }                      // if specified
            public string Operator { get; set; } = ">=";         // default like "at least 1"
            public Dictionary<string, string>? Attributes { get; set; }
            public Dictionary<string, JToken>? Properties { get; set; }
        }

        private static int ParseSleepDurationMs(TestStep step)
        {
            string? rawDuration = string.IsNullOrWhiteSpace(step.Value) ? step.Target : step.Value;

            if (string.IsNullOrWhiteSpace(rawDuration))
                throw new ArgumentException("sleep command requires a duration in the Value column, for example '1500ms', '2s', or '00:00:05'.");

            string text = rawDuration.Trim();

            if (TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out var parsedTimeSpan))
            {
                return ValidateSleepDuration(parsedTimeSpan.TotalMilliseconds, text);
            }

            string normalized = text.Replace(" ", string.Empty);

            if (normalized.EndsWith("ms", StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(normalized[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out double milliseconds))
            {
                return ValidateSleepDuration(milliseconds, text);
            }

            if (normalized.EndsWith("s", StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(normalized[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double seconds))
            {
                return ValidateSleepDuration(TimeSpan.FromSeconds(seconds).TotalMilliseconds, text);
            }

            if (normalized.EndsWith("m", StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(normalized[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double minutes))
            {
                return ValidateSleepDuration(TimeSpan.FromMinutes(minutes).TotalMilliseconds, text);
            }

            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double plainNumber))
            {
                return ValidateSleepDuration(TimeSpan.FromSeconds(plainNumber).TotalMilliseconds, text);
            }

            throw new ArgumentException($"Invalid sleep duration '{text}'. Use formats like '1500ms', '2s', '1.5m', or '00:00:05'.");
        }

        private static int ValidateSleepDuration(double totalMilliseconds, string rawValue)
        {
            if (double.IsNaN(totalMilliseconds) || double.IsInfinity(totalMilliseconds) || totalMilliseconds < 0)
                throw new ArgumentOutOfRangeException(nameof(rawValue), rawValue, "sleep duration must be zero or greater.");

            if (totalMilliseconds > int.MaxValue)
                throw new ArgumentOutOfRangeException(nameof(rawValue), rawValue, $"sleep duration cannot exceed {int.MaxValue} ms.");

            return (int)Math.Round(totalMilliseconds, MidpointRounding.AwayFromZero);
        }

        private void WaitForElementAssertion(TestStep s)
        {
            var selectors = s.Selectors ?? new List<string[]> { new[] { s.Target! } };
            var cfg = ParseWaitForElementCfg(s.Value);

            var wait = new WebDriverWait(_driver!, TimeSpan.FromMilliseconds(cfg.TimeoutMs));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(StaleElementReferenceException), typeof(WebDriverException));

            bool ok = wait.Until(_ => TrySatisfyWaitForElement(selectors, cfg));
            if (!ok)
                throw new WebDriverTimeoutException("waitForElement timed out.");
        }
        private void DispatchLightningInputAndChange(IWebElement element)
        {
            try
            {
                var js = (IJavaScriptExecutor)_driver!;
                const string script = @"
            const el = arguments[0];
            try {
                el.dispatchEvent(new Event('input', { bubbles: true }));
            } catch (e) {}
            try {
                el.dispatchEvent(new Event('change', { bubbles: true }));
            } catch (e) {}
        ";
                js.ExecuteScript(script, element);
            }
            catch
            {
                // If JS fails, do nothing (we never want this to break CRM stability)
            }
        }
        private static string NormalizeSelectValueToText(string raw)
        {
            raw = (raw ?? "").Trim();

            if (raw.StartsWith("label=", StringComparison.OrdinalIgnoreCase))
                return raw["label=".Length..].Trim();

            if (raw.StartsWith("value=", StringComparison.OrdinalIgnoreCase))
                return raw["value=".Length..].Trim();

            // For Lightning, index-based selection isn’t reliable (order changes).
            // We’ll still return raw; user should provide text for Lightning.
            if (raw.StartsWith("index=", StringComparison.OrdinalIgnoreCase))
                return raw; // will likely fail unless you handle index mode separately

            return raw;
        }

        private void SelectLightningComboboxOption(IWebElement target, string optionText)
        {
            // Lightning targets vary: sometimes the recorder points to the input, sometimes the button.
            // We open the dropdown from whichever element we received.

            // 1) click the target (or nearest combobox button)
            TryOpenLightningCombobox(target);

            // 2) wait for listbox, then click exact option text
            var wait = new WebDriverWait(_driver!, TimeSpan.FromSeconds(10));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(StaleElementReferenceException));

            // Find option by exact normalized text (stable)
            string optionXpath = $"//div[@role='listbox']//*[normalize-space(.)={QuoteForXpath(optionText)}]";

            IWebElement option = wait.Until(d =>
            {
                try
                {
                    // use driver so it can search in the current frame context
                    var el = d.FindElement(By.XPath(optionXpath));
                    return el.Displayed ? el : null!;
                }
                catch { return null!; }
            });

            option.Click();
        }

        private void TryOpenLightningCombobox(IWebElement target)
        {
            try
            {
                // If the element itself is a combobox/button-like, click it
                string role = "";
                try { role = target.GetAttribute("role") ?? ""; } catch { }

                if (role.Equals("combobox", StringComparison.OrdinalIgnoreCase) ||
                    role.Equals("button", StringComparison.OrdinalIgnoreCase))
                {
                    target.Click();
                    return;
                }

                // Try nearest combobox button from container
                // Common lightning structure: button[role='combobox'] or button.slds-combobox__input
                var container = target.FindElement(By.XPath("./ancestor::*[contains(@class,'slds-combobox')][1]"));
                var btns = container.FindElements(By.XPath(".//button[@role='combobox' or contains(@class,'slds-combobox__input')]"));
                if (btns.Count > 0)
                {
                    btns[0].Click();
                    return;
                }

                // Fallback: click target anyway
                target.Click();
            }
            catch
            {
                // last resort: click target (may throw, that’s ok)
                target.Click();
            }
        }


        private WaitForElementCfg ParseWaitForElementCfg(string? json)
        {
            var cfg = new WaitForElementCfg();

            if (string.IsNullOrWhiteSpace(json))
                return cfg;

            try
            {
                var jo = JObject.Parse(json);

                if (jo.TryGetValue("timeoutMs", out var t) && t.Type == JTokenType.Integer)
                    cfg.TimeoutMs = t.Value<int>();

                if (jo.TryGetValue("visible", out var v) && v.Type == JTokenType.Boolean)
                    cfg.Visible = v.Value<bool>();

                if (jo.TryGetValue("count", out var c) && c.Type == JTokenType.Integer)
                    cfg.Count = c.Value<int>();

                if (jo.TryGetValue("operator", out var op) && op.Type == JTokenType.String)
                    cfg.Operator = op.Value<string>()?.Trim() ?? cfg.Operator;

                if (jo["attributes"] is JObject attrs && attrs.HasValues)
                    cfg.Attributes = attrs.Properties().ToDictionary(p => p.Name, p => p.Value.ToString());

                if (jo["properties"] is JObject props && props.HasValues)
                    cfg.Properties = props.Properties().ToDictionary(p => p.Name, p => p.Value);
            }
            catch
            {
                //If Value isn't valid JSON, just treat it as "exists/visible" with default timeout.
            }

            return cfg;
        }

        private bool TrySatisfyWaitForElement(List<string[]> selectors, WaitForElementCfg cfg)
        {
            //Try to find *visible* elements (matches your “visually apparent” need)
            if (!TryLocateAllDisplayed(selectors, out var elements))
            {
                //If the assertion is explicitly "not visible", then "can't find visible elements" is success.
                if (cfg.Visible == false) return true;

                //If user asked for count==0 (or similar), treat "0 visible" as the actual count.
                if (cfg.Count.HasValue && CompareCount(0, cfg.Count.Value, cfg.Operator)) return true;

                return false;
            }

            int actualCount = elements.Count;

            //Count/operator assertion
            if (cfg.Count.HasValue && !CompareCount(actualCount, cfg.Count.Value, cfg.Operator))
                return false;

            //Visible false means "no visible matches"
            if (cfg.Visible == false && actualCount > 0)
                return false;

            //Attributes (any matching element satisfies)
            if (cfg.Attributes != null && cfg.Attributes.Count > 0)
            {
                bool any = elements.Any(e =>
                    cfg.Attributes.All(kv =>
                        string.Equals(e.GetAttribute(kv.Key), kv.Value, StringComparison.Ordinal)
                    )
                );
                if (!any) return false;
            }

            //Properties via JS (any matching element satisfies)
            if (cfg.Properties != null && cfg.Properties.Count > 0)
            {
                var js = (IJavaScriptExecutor)_driver!;
                bool any = elements.Any(e => cfg.Properties.All(kv => PropertyEquals(js, e, kv.Key, kv.Value)));
                if (!any) return false;
            }

            return true;
        }

        private static bool CompareCount(int actual, int expected, string? op)
        {
            op = (op ?? ">=").Trim();

            return op switch
            {
                "==" or "=" => actual == expected,
                "!=" => actual != expected,
                ">" => actual > expected,
                ">=" => actual >= expected,
                "<" => actual < expected,
                "<=" => actual <= expected,
                _ => actual >= expected
            };
        }

        private static bool PropertyEquals(IJavaScriptExecutor js, IWebElement el, string propKey, JToken expected)
        {
            string path = propKey.StartsWith(".") ? propKey : "." + propKey;

            object? actual;
            try
            {
                actual = js.ExecuteScript($"return arguments[0]{path};", el);
            }
            catch
            {
                return false;
            }

            //Compare by expected token type (basic support)
            try
            {
                return expected.Type switch
                {
                    JTokenType.Boolean => actual is bool b && b == expected.Value<bool>(),
                    JTokenType.Integer => Convert.ToInt64(actual) == expected.Value<long>(),
                    JTokenType.Float => Math.Abs(Convert.ToDouble(actual) - expected.Value<double>()) < 0.000001,
                    _ => string.Equals(Convert.ToString(actual) ?? "", expected.ToString(), StringComparison.Ordinal)
                };
            }
            catch
            {
                return false;
            }
        }


        private const string DataFileTokenPrefix = "${DATAFILE:";
        private const string DataFileTokenSuffix = "}";

        private static bool TryParseDataFileToken(string? value, out string fileName)
        {
            fileName = "";
            if (string.IsNullOrWhiteSpace(value)) return false;

            value = value.Trim();
            if (!value.StartsWith(DataFileTokenPrefix, StringComparison.OrdinalIgnoreCase)) return false;
            if (!value.EndsWith(DataFileTokenSuffix, StringComparison.Ordinal)) return false;

            fileName = value.Substring(DataFileTokenPrefix.Length, value.Length - DataFileTokenPrefix.Length - DataFileTokenSuffix.Length).Trim();
            return !string.IsNullOrWhiteSpace(fileName);
        }

        private static bool TryExtractFakePathFileName(string? rawValue, out string fileName)
        {
            fileName = "";
            if (string.IsNullOrWhiteSpace(rawValue)) return false;
            if (rawValue.IndexOf("fakepath", StringComparison.OrdinalIgnoreCase) < 0) return false;

            fileName = Path.GetFileName(rawValue.Trim());
            return !string.IsNullOrWhiteSpace(fileName);
        }

        private string ResolveDataFileAbsolutePath(string fileName)
        {
            if (string.IsNullOrWhiteSpace(_dataFilesDir))
                throw new InvalidOperationException("DataFiles directory not initialized.");

            var full = Path.GetFullPath(Path.Combine(_dataFilesDir, fileName));

            if (!File.Exists(full))
                throw new FileNotFoundException($"Data file not found: {full}");

            return full;
        }
        private string? TryResolveUploadPath(string? rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return null;

            rawValue = rawValue.Trim();

            // ${DATAFILE:filename.ext}
            if (TryParseDataFileToken(rawValue, out var tokenFile))
                return ResolveDataFileAbsolutePath(tokenFile);

            // C:\fakepath\filename.ext
            if (TryExtractFakePathFileName(rawValue, out var fakeFile))
                return ResolveDataFileAbsolutePath(fakeFile);

            return null;
        }



        // ===== Locator fallback with fast CRM frame hopping =====

        // ===== Locator fallback with fast CRM frame hopping =====

        private IWebElement Locate(List<string[]> selectors)
        {
            bool isSalesforceMode = ShouldUseSalesforceMode(selectors);
            bool isCrm = IsCrmContext();


            if (!isCrm && !isSalesforceMode)
            {
                return LocateFastNoFrames(selectors);
            }


            foreach (var selectorGroup in selectors)
            {
                //Salesforce-only: re-rank fallback order inside this selector group
                IEnumerable<string> orderedSelectors = selectorGroup;

                if (isSalesforceMode)
                    orderedSelectors = selectorGroup.OrderBy(sel => SelectorRank(sel));

                foreach (var selRaw in orderedSelectors)
                {
                    var sel = (selRaw ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(sel))
                        continue;

                    //Salesforce: handle pierce/ + aria/ specially (do NOT break CRM behavior)
                    if (isSalesforceMode && sel.StartsWith("pierce/", StringComparison.OrdinalIgnoreCase))
                    {
                        string css = sel.Substring("pierce/".Length).Trim();

                        //1) current context
                        if (TryFindPierceInCurrentContext(css, out var found, TimeSpan.FromMilliseconds(700)))
                            return found;

                        //2) last good frame
                        if (_lastFrameCss != null)
                        {
                            if (TrySwitchToFrameCss(_lastFrameCss, TimeSpan.FromMilliseconds(400)) &&
                                TryFindPierceInCurrentContext(css, out found, TimeSpan.FromMilliseconds(700)))
                                return found;

                            _driver!.SwitchTo().DefaultContent();
                        }

                        //3) known CRM frames
                        if (TryFindPierceAcrossKnownCrmFrames(css, out found))
                            return found;

                        //4) quick scan of visible iframes
                        if (TryFindPierceAcrossAllIframes(css, out found))
                            return found;

                        _driver!.SwitchTo().DefaultContent();
                        continue;
                    }

                    if (isSalesforceMode && sel.StartsWith("aria/", StringComparison.OrdinalIgnoreCase))
                    {
                        var ariaBys = BuildAriaCandidateBys(sel);

                        foreach (var by in ariaBys)
                        {
                            //1) current context
                            if (TryFindInCurrentContext(by, out var found, TimeSpan.FromMilliseconds(700)))
                                return found;

                            //2) last good frame
                            if (_lastFrameCss != null)
                            {
                                if (TrySwitchToFrameCss(_lastFrameCss, TimeSpan.FromMilliseconds(400)) &&
                                    TryFindInCurrentContext(by, out found, TimeSpan.FromMilliseconds(700)))
                                    return found;

                                _driver!.SwitchTo().DefaultContent();
                            }

                            //3) known CRM frames
                            if (TryFindAcrossKnownCrmFrames(by, out found))
                                return found;

                            //4) quick scan of visible iframes
                            if (TryFindAcrossAllIframes(by, out found))
                                return found;

                            _driver!.SwitchTo().DefaultContent();
                        }

                        continue;
                    }

                    //Default path (CRM-safe)
                    var byDefault = ToBy(sel);

                    //1) current context
                    if (TryFindInCurrentContext(byDefault, out var foundDefault, TimeSpan.FromMilliseconds(700)))
                        return foundDefault;

                    //2) last good frame
                    if (_lastFrameCss != null)
                    {
                        if (TrySwitchToFrameCss(_lastFrameCss, TimeSpan.FromMilliseconds(400)) &&
                            TryFindInCurrentContext(byDefault, out foundDefault, TimeSpan.FromMilliseconds(700)))
                            return foundDefault;

                        _driver!.SwitchTo().DefaultContent();
                    }

                    //3) known CRM frames
                    if (TryFindAcrossKnownCrmFrames(byDefault, out foundDefault))
                        return foundDefault;

                    //4) quick scan of visible iframes
                    if (TryFindAcrossAllIframes(byDefault, out foundDefault))
                        return foundDefault;

                    _driver!.SwitchTo().DefaultContent();
                }
            }

            throw new NoSuchElementException("None of the provided selectors were found.");
        }

        private bool ShouldUseSalesforceMode(List<string[]> selectors)
        {
            try
            {
                //Strong signals for Lightning-style DOM
                bool hasPierce = selectors.Any(g => g.Any(s =>
                    !string.IsNullOrWhiteSpace(s) && s.Trim().StartsWith("pierce/", StringComparison.OrdinalIgnoreCase)));

                if (hasPierce) return true;

                //Dynamic Lightning-ish IDs (input-### / combobox-button-###)
                bool hasDynamic = selectors.Any(g => g.Any(s =>
                {
                    s = (s ?? "").Trim();
                    return s.Contains("input-", StringComparison.OrdinalIgnoreCase) ||
                           s.Contains("combobox-button-", StringComparison.OrdinalIgnoreCase) ||
                           s.Contains("help-message-", StringComparison.OrdinalIgnoreCase);
                }));

                if (hasDynamic) return true;

                //URL heuristic (only if applicable)
                var url = _driver?.Url ?? "";
                if (url.Contains("/lightning/", StringComparison.OrdinalIgnoreCase) ||
                    url.Contains("lightning.force.com", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { /* ignore */ }

            //Default: CRM/Portal mode
            return false;
        }

        private bool TryLocateHtmlSelect(List<string[]> selectors, out IWebElement selectEl)
        {
            selectEl = null!;

            // Prefer selectors that can point to controls (avoid aria/text/pierce)
            var flat = selectors.SelectMany(g => g)
            .Select(s => (s ?? "").Trim())
            .Where(s => s.Length > 0)
            .Where(s =>
            s.StartsWith("css=", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("id=", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("name=", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("//") || s.StartsWith("/"))
            .ToList();

            foreach (var sel in flat)
            {
                var by = ToBy(sel);

                // current context
                if (TryFindSelectInCurrentContext(by, out selectEl, TimeSpan.FromMilliseconds(400)))
                    return true;

                // last known frame
                if (_lastFrameCss != null)
                {
                    if (TrySwitchToFrameCss(_lastFrameCss, TimeSpan.FromMilliseconds(250)) &&
                    TryFindSelectInCurrentContext(by, out selectEl, TimeSpan.FromMilliseconds(400)))
                        return true;

                    _driver!.SwitchTo().DefaultContent();
                }

                // known CRM frames
                if (TryFindSelectAcrossKnownCrmFrames(by, out selectEl))
                    return true;

                // scan visible iframes
                if (TryFindSelectAcrossAllIframes(by, out selectEl))
                    return true;

                _driver!.SwitchTo().DefaultContent();
            }

            return false;
        }

        private bool TryFindSelectInCurrentContext(By by, out IWebElement element, TimeSpan timeout)
        {
            var end = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < end)
            {
                try
                {
                    var els = _driver!.FindElements(by);
                    foreach (var el in els)
                    {
                        try
                        {
                            if (string.Equals(el.TagName, "select", StringComparison.OrdinalIgnoreCase))
                            {
                                element = el;
                                return true;
                            }
                        }
                        catch { /* ignore */ }
                    }
                }
                catch { /* ignore */ }

                System.Threading.Thread.Sleep(70);
            }

            element = null!;
            return false;
        }

        private bool TryFindSelectAcrossKnownCrmFrames(By by, out IWebElement element)
        {
            var known = new[]
            {
"iframe#AppLandingPage",
"iframe[id*='contentIFrame']",
"iframe[title*='App Landing']",
"iframe[title*='Content']",
"iframe[data-id='contentIFrame0']",
"iframe[name='contentIFrame0']"
};

            foreach (var css in known)
            {
                if (TrySwitchToFrameCss(css, TimeSpan.FromMilliseconds(300)))
                {
                    if (TryFindSelectInCurrentContext(by, out element, TimeSpan.FromMilliseconds(450)))
                    {
                        _lastFrameCss = css;
                        return true;
                    }
                    _driver!.SwitchTo().DefaultContent();
                }
            }

            element = null!;
            return false;
        }

        private bool TryFindSelectAcrossAllIframes(By by, out IWebElement element)
        {
            try
            {
                _driver!.SwitchTo().DefaultContent();
                var frames = _driver.FindElements(By.CssSelector("iframe"));

                foreach (var f in frames)
                {
                    if (!SafeDisplayed(f)) continue;

                    _driver.SwitchTo().Frame(f);

                    if (TryFindSelectInCurrentContext(by, out element, TimeSpan.FromMilliseconds(350)))
                        return true;

                    _driver.SwitchTo().DefaultContent();
                }
            }
            catch { /* ignore */ }

            element = null!;
            return false;
        }

        private void SelectOnHtmlSelect(IWebElement selectEl, string rawVal)
        {
            rawVal = (rawVal ?? "").Trim();

            string? desiredValue = null;
            string? desiredText = null;

            if (rawVal.StartsWith("value=", StringComparison.OrdinalIgnoreCase))
                desiredValue = rawVal["value=".Length..].Trim();
            else if (rawVal.StartsWith("label=", StringComparison.OrdinalIgnoreCase))
                desiredText = rawVal["label=".Length..].Trim();
            else if (rawVal.All(char.IsDigit))
                desiredValue = rawVal; // CRM option-set numeric values
            else
                desiredText = rawVal;

            // Try standard SelectElement
            try
            {
                var select = new SelectElement(selectEl);

                if (!string.IsNullOrWhiteSpace(desiredValue))
                {
                    select.SelectByValue(desiredValue);
                    return;
                }

                if (!string.IsNullOrWhiteSpace(desiredText))
                {
                    select.SelectByText(desiredText);
                    return;
                }
            }
            catch
            {
                // hidden/non-interactable selects often throw
            }

            // JS fallback: set value + fire events
            string valueToSet = desiredValue ?? FindOptionValueByText(selectEl, desiredText ?? "");
            if (string.IsNullOrWhiteSpace(valueToSet))
                throw new Exception($"Could not resolve select option for '{rawVal}'");

            var js = (IJavaScriptExecutor)_driver!;
            js.ExecuteScript(@"
const sel = arguments[0];
const val = arguments[1];
sel.value = val;
sel.dispatchEvent(new Event('input', { bubbles: true }));
sel.dispatchEvent(new Event('change', { bubbles: true }));
", selectEl, valueToSet);
        }

        private string FindOptionValueByText(IWebElement selectEl, string text)
        {
            text = (text ?? "").Trim();
            if (text.Length == 0) return "";

            try
            {
                var opts = selectEl.FindElements(By.TagName("option"));
                foreach (var opt in opts)
                {
                    try
                    {
                        if (string.Equals(opt.Text?.Trim(), text, StringComparison.OrdinalIgnoreCase))
                            return opt.GetAttribute("value") ?? "";
                    }
                    catch { /* ignore */ }
                }
            }
            catch { /* ignore */ }

            return "";
        }
        private int SelectorRank(string sel)
        {
            sel = (sel ?? "").Trim();

            // Best first
            if (sel.StartsWith("pierce/", StringComparison.OrdinalIgnoreCase)) return 0;
            if (sel.StartsWith("aria/", StringComparison.OrdinalIgnoreCase)) return 1;

            // Stable selectors
            if (sel.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase)) return IsDynamicIdSelector(sel) ? 90 : 10;
            if (sel.StartsWith("css=", StringComparison.OrdinalIgnoreCase)) return IsDynamicIdSelector(sel) ? 90 : 11;
            if (sel.StartsWith("name=", StringComparison.OrdinalIgnoreCase)) return 12;
            if (sel.StartsWith("id=", StringComparison.OrdinalIgnoreCase)) return IsDynamicIdSelector(sel) ? 95 : 13;

            // Raw CSS/XPath
            if (sel.StartsWith("//") || sel.StartsWith("/")) return IsDynamicIdSelector(sel) ? 90 : 15;

            // Raw css selector (often includes #id)
            return IsDynamicIdSelector(sel) ? 95 : 20;
        }

        private bool IsDynamicIdSelector(string sel)
        {
            sel = (sel ?? "").Trim();

            // Common Salesforce dynamic IDs
            // #input-203, #combobox-button-508, #help-message-123, etc.
            // Also handles css=#input-203 and xpath=//*[@id='input-203']
            string[] patterns =
            {
        "input-",
        "combobox-button-",
        "help-message-",
        "dropdown-element-",
        "brandBand_",
        "split-left-",
    };

            // quick checks
            foreach (var p in patterns)
            {
                if (sel.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // If it looks like "...<pattern><digits>" treat as dynamic
                    // Simple numeric sniff:
                    int idx = sel.IndexOf(p, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0)
                    {
                        string tail = sel.Substring(idx + p.Length);
                        if (tail.Any(char.IsDigit)) return true;
                    }
                }
            }

            return false;
        }
        private IWebElement? DeepQuerySelector(string cssSelector)
        {
            var js = (IJavaScriptExecutor)_driver!;
            const string script = @"
        const selector = arguments[0];

        function* walk(root) {
            yield root;
            const tree = root.querySelectorAll ? root.querySelectorAll('*') : [];
            for (const el of tree) yield el;
        }

        function deepQuery(root, selector) {
            // Try normal querySelector first
            if (root.querySelector) {
                const found = root.querySelector(selector);
                if (found) return found;
            }

            // Walk and traverse shadow roots
            for (const el of walk(root)) {
                if (el && el.shadowRoot) {
                    const found = deepQuery(el.shadowRoot, selector);
                    if (found) return found;
                }
            }
            return null;
        }

        return deepQuery(document, selector);
    ";

            var found = js.ExecuteScript(script, cssSelector);
            return found as IWebElement;
        }

        private bool TryFindPierceInCurrentContext(string css, out IWebElement element, TimeSpan timeout)
        {
            var end = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < end)
            {
                try
                {
                    var found = DeepQuerySelector(css);
                    if (found != null && SafeDisplayed(found))
                    {
                        element = found;
                        return true;
                    }
                }
                catch { /* ignore */ }

                System.Threading.Thread.Sleep(120);
            }

            element = null!;
            return false;
        }
        private List<By> BuildAriaCandidateBys(string ariaSelector)
        {
            // Expected formats:
            // aria/Description[role="textbox"]
            // aria/Description
            string raw = ariaSelector.Trim();
            string body = raw.Substring("aria/".Length);

            string labelText = body;
            string? role = null;

            int bracket = body.IndexOf('[');
            if (bracket >= 0)
            {
                labelText = body.Substring(0, bracket).Trim();

                int roleIdx = body.IndexOf("role=", StringComparison.OrdinalIgnoreCase);
                if (roleIdx >= 0)
                {
                    // capture role value inside quotes if present
                    int quote1 = body.IndexOf('"', roleIdx);
                    int quote2 = quote1 >= 0 ? body.IndexOf('"', quote1 + 1) : -1;
                    if (quote1 >= 0 && quote2 > quote1)
                        role = body.Substring(quote1 + 1, quote2 - quote1 - 1).Trim();
                }
            }
            else
            {
                labelText = labelText.Trim();
            }

            // Candidate 1: aria-label exact
            // If role exists, include it; otherwise just aria-label
            string css1 = role != null
                ? $"*[aria-label='{EscapeCssAttr(labelText)}'][role='{EscapeCssAttr(role)}']"
                : $"*[aria-label='{EscapeCssAttr(labelText)}']";

            // Candidate 2: label/container strategy (Salesforce slds-form-element)
            // Find label text then search within the form element container for input/textarea
            // Use both exact and contains to be resilient.
            string labelXpathExact =
                $"(//label[normalize-space(.)={QuoteForXpath(labelText)}] | " +
                $"//span[normalize-space(.)={QuoteForXpath(labelText)}] | " +
                $"//*[@title={QuoteForXpath(labelText)}])";

            string inputXpath =
                role != null && role.Equals("textbox", StringComparison.OrdinalIgnoreCase)
                    ? ".//textarea | .//input"
                    : ".//textarea | .//input | .//*[@role='textbox']";

            string sldsContainer =
                $"{labelXpathExact}/ancestor::*[contains(@class,'slds-form-element')][1]{inputXpath}";

            // Candidate 3: broader label contains fallback
            string labelXpathContains =
                $"(//label[contains(normalize-space(.), {QuoteForXpath(labelText)})] | " +
                $"//span[contains(normalize-space(.), {QuoteForXpath(labelText)})])" +
                $"/ancestor::*[contains(@class,'slds-form-element')][1]{inputXpath}";

            return new List<By>
    {
        By.CssSelector(css1),
        By.XPath(sldsContainer),
        By.XPath(labelXpathContains)
    };
        }

        private static string EscapeCssAttr(string s)
        {
            // minimal escaping for single quotes in attribute values
            return (s ?? "").Replace("'", "\\'");
        }

        private bool TryFindPierceAcrossKnownCrmFrames(string css, out IWebElement element)
        {
            var known = new[]
            {
        "iframe#AppLandingPage",
        "iframe[id*='contentIFrame']",
        "iframe[title*='App Landing']",
        "iframe[title*='Content']",
        "iframe[data-id='contentIFrame0']",
        "iframe[name='contentIFrame0']"
    };

            foreach (var frameCss in known)
            {
                if (TrySwitchToFrameCss(frameCss, TimeSpan.FromMilliseconds(400)))
                {
                    if (TryFindPierceInCurrentContext(css, out element, TimeSpan.FromMilliseconds(800)))
                    {
                        _lastFrameCss = frameCss;
                        return true;
                    }
                    _driver!.SwitchTo().DefaultContent();
                }
            }

            element = null!;
            return false;
        }

        private bool TryFindPierceAcrossAllIframes(string css, out IWebElement element)
        {
            try
            {
                _driver!.SwitchTo().DefaultContent();
                var frames = _driver.FindElements(By.CssSelector("iframe"));
                foreach (var f in frames)
                {
                    if (!SafeDisplayed(f)) continue;

                    _driver.SwitchTo().Frame(f);

                    if (TryFindPierceInCurrentContext(css, out element, TimeSpan.FromMilliseconds(500)))
                        return true;

                    _driver.SwitchTo().DefaultContent();
                }
            }
            catch { /* ignore */ }

            element = null!;
            return false;
        }

        private By ToBy(string selector)
        {
            selector = (selector ?? "").Trim();

            if (selector.StartsWith("id=", StringComparison.OrdinalIgnoreCase)) return By.Id(selector[3..]);
            if (selector.StartsWith("css=", StringComparison.OrdinalIgnoreCase)) return By.CssSelector(selector[4..]);
            if (selector.StartsWith("name=", StringComparison.OrdinalIgnoreCase)) return By.Name(selector[5..]);

            // ---- XPATH (fix legacy bad formats) ----
            if (selector.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase))
            {
                var xp = selector.Substring("xpath=".Length).Trim();

                // Legacy Excel + some conversions produce: *[@id="x"]  (missing //)
                if (xp.StartsWith("*", StringComparison.Ordinal))
                    xp = "//" + xp;

                // If it doesn't start with '/', '.', '(' and also not with '//' then assume it needs '//'
                if (!xp.StartsWith("/", StringComparison.Ordinal) &&
                    !xp.StartsWith(".", StringComparison.Ordinal) &&
                    !xp.StartsWith("(", StringComparison.Ordinal) &&
                    !xp.StartsWith("//", StringComparison.Ordinal))
                {
                    xp = "//" + xp;
                }

                return By.XPath(xp);
            }

            // Bare xpath
            if (selector.StartsWith("//", StringComparison.Ordinal) || selector.StartsWith("/", StringComparison.Ordinal))
                return By.XPath(selector);

            // ---- ARIA (use XPath instead of CSS to avoid quote/bracket issues) ----
            if (selector.StartsWith("aria/", StringComparison.OrdinalIgnoreCase))
            {
                // Supports:
                // aria/LabelText
                // aria/LabelText[role="textbox"]
                // aria/[role="listbox"]
                var raw = selector.Substring("aria/".Length).Trim();

                string label = raw;
                string? role = null;

                int bracket = raw.IndexOf('[');
                if (bracket >= 0)
                {
                    label = raw.Substring(0, bracket).Trim();

                    // parse role="..."
                    var m = System.Text.RegularExpressions.Regex.Match(raw, @"role\s*=\s*""([^""]+)""");
                    if (m.Success) role = m.Groups[1].Value.Trim();
                }

                // aria/[role="listbox"] (label is empty)
                if (string.IsNullOrWhiteSpace(label) && !string.IsNullOrWhiteSpace(role))
                    return By.XPath($"//*[@role={QuoteForXpath(role)}]");

                if (!string.IsNullOrWhiteSpace(label) && !string.IsNullOrWhiteSpace(role))
                    return By.XPath($"//*[@aria-label={QuoteForXpath(label)} and @role={QuoteForXpath(role)}]");

                if (!string.IsNullOrWhiteSpace(label))
                    return By.XPath($"//*[@aria-label={QuoteForXpath(label)}]");

                // If totally empty, fall back (won’t usually happen)
                return By.XPath("//*");
            }

            if (selector.StartsWith("text=", StringComparison.OrdinalIgnoreCase))
                return By.XPath("//*[contains(normalize-space(.), " + QuoteForXpath(selector[5..]) + ")]");

            // fallback: treat as CSS
            return By.CssSelector(selector);
        }



        private static string QuoteForXpath(string text)
        {
            if (!text.Contains("'")) return $"'{text}'";
            if (!text.Contains("\"")) return $"\"{text}\"";
            var parts = text.Split('\'');
            return "concat(" + string.Join(", \"'\", ", parts.Select(p => $"'{p}'")) + ")";
        }

        private bool TryFindInCurrentContext(By by, out IWebElement element, TimeSpan timeout)
        {
            var end = DateTime.UtcNow + timeout;
            Exception? last = null;
            while (DateTime.UtcNow < end)
            {
                try
                {
                    var el = _driver!.FindElement(by);
                    if (el.Displayed)
                    {
                        element = el;
                        return true;
                    }
                }
                catch (Exception ex) { last = ex; }
                System.Threading.Thread.Sleep(120);
            }
            element = null!;
            return false;
        }

        private bool TrySwitchToFrameCss(string css, TimeSpan timeout)
        {
            var end = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < end)
            {
                try
                {
                    _driver!.SwitchTo().DefaultContent();
                    var frame = _driver.FindElement(By.CssSelector(css));
                    _driver.SwitchTo().Frame(frame);
                    return true;
                }
                catch { System.Threading.Thread.Sleep(100); }
            }
            return false;
        }

        private bool TryFindAcrossKnownCrmFrames(By by, out IWebElement element)
        {
            var known = new[]
            {
                "iframe#AppLandingPage",
                "iframe[id*='contentIFrame']",
                "iframe[title*='App Landing']",
                "iframe[title*='Content']",
                "iframe[data-id='contentIFrame0']",
                "iframe[name='contentIFrame0']"
            };

            foreach (var css in known)
            {
                if (TrySwitchToFrameCss(css, TimeSpan.FromMilliseconds(400)))
                {
                    if (TryFindInCurrentContext(by, out element, TimeSpan.FromMilliseconds(800)))
                    {
                        _lastFrameCss = css;
                        return true;
                    }
                    _driver!.SwitchTo().DefaultContent();
                }
            }
            element = null!;
            return false;
        }

        private bool TryFindAcrossAllIframes(By by, out IWebElement element)
        {
            try
            {
                _driver!.SwitchTo().DefaultContent();
                var frames = _driver.FindElements(By.CssSelector("iframe"));
                foreach (var f in frames)
                {
                    if (!f.Displayed) continue;
                    _driver.SwitchTo().Frame(f);
                    if (TryFindInCurrentContext(by, out element, TimeSpan.FromMilliseconds(500)))
                        return true;
                    _driver.SwitchTo().DefaultContent();
                }
            }
            catch { /* ignore */ }

            element = null!;
            return false;
        }

        private void TryEnterCrmMainFrame()
        {
            if (_driver == null) return;

            _driver.SwitchTo().DefaultContent();

            var candidateIds = new[]
            {
                "contentIFrame0","contentIFrame1","contentIFrame",
                "AppLandingPage","dashboardFrame","InlineDialog_Iframe","InlineDialog",
                "fullPageWebResource","WebResource_"
            };

            foreach (var id in candidateIds)
            {
                var byId = _driver.FindElements(By.CssSelector($"iframe#{id}, iframe[id^='{id}']"));
                if (byId.Count > 0)
                {
                    _driver.SwitchTo().Frame(byId[0]);
                    return;
                }
            }
            /*
                        var frames = _driver.FindElements(By.TagName("iframe"));
                        if (frames.Count > 0)
                            _driver.SwitchTo().Frame(frames[0]); */
        }

        private bool TryLocateAllDisplayed(List<string[]> selectors, out List<IWebElement> elements)
        {
            elements = new List<IWebElement>();

            foreach (var selectorGroup in selectors)
            {
                foreach (var sel in selectorGroup)
                {
                    var by = ToBy(sel);

                    // 1) current context
                    if (TryFindAllInCurrentContext(by, out elements, TimeSpan.FromMilliseconds(300)))
                        return true;

                    // 2) last good frame
                    if (_lastFrameCss != null)
                    {
                        if (TrySwitchToFrameCss(_lastFrameCss, TimeSpan.FromMilliseconds(300)) &&
                            TryFindAllInCurrentContext(by, out elements, TimeSpan.FromMilliseconds(300)))
                            return true;

                        _driver!.SwitchTo().DefaultContent();
                    }

                    // 3) known CRM frames
                    if (TryFindAllAcrossKnownCrmFrames(by, out elements))
                        return true;

                    // 4) quick scan of visible iframes
                    if (TryFindAllAcrossAllIframes(by, out elements))
                        return true;

                    _driver!.SwitchTo().DefaultContent();
                }
            }

            return false;
        }

        private bool TryFindAllInCurrentContext(By by, out List<IWebElement> elements, TimeSpan timeout)
        {
            var end = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < end)
            {
                try
                {
                    var found = _driver!.FindElements(by)
                        .Where(SafeDisplayed)
                        .ToList();

                    if (found.Count > 0)
                    {
                        elements = found;
                        return true;
                    }
                }
                catch { /* ignore */ }

                System.Threading.Thread.Sleep(120);
            }

            elements = new List<IWebElement>();
            return false;
        }

        private bool TryFindAllAcrossKnownCrmFrames(By by, out List<IWebElement> elements)
        {
            var known = new[]
            {
        "iframe#AppLandingPage",
        "iframe[id*='contentIFrame']",
        "iframe[title*='App Landing']",
        "iframe[title*='Content']",
        "iframe[data-id='contentIFrame0']",
        "iframe[name='contentIFrame0']"
    };

            foreach (var css in known)
            {
                if (TrySwitchToFrameCss(css, TimeSpan.FromMilliseconds(400)))
                {
                    if (TryFindAllInCurrentContext(by, out elements, TimeSpan.FromMilliseconds(500)))
                    {
                        _lastFrameCss = css;
                        return true;
                    }
                    _driver!.SwitchTo().DefaultContent();
                }
            }

            elements = new List<IWebElement>();
            return false;
        }

        private IWebElement LocateFastNoFrames(List<string[]> selectors)
        {
            // No frame hopping, just try in current document quickly
            foreach (var selectorGroup in selectors)
            {
                foreach (var selRaw in selectorGroup)
                {
                    var sel = (selRaw ?? "").Trim();
                    if (sel.Length == 0) continue;

                    var by = ToBy(sel);

                    // small timeout (portal pages should be direct + fast)
                    if (TryFindInCurrentContext(by, out var found, TimeSpan.FromMilliseconds(250)))
                        return found;
                }
            }

            throw new NoSuchElementException("None of the provided selectors were found (fast path).");
        }

        private void LightningSafeClick(IWebElement el)
        {
            try
            {
                ScrollIntoView(el);
                el.Click();
                return;
            }
            catch (ElementClickInterceptedException)
            {
                // common in Lightning dialogs
            }
            catch (WebDriverException)
            {
                // fallback below
            }

            // JS click fallback (Lightning overlay situations)
            try
            {
                ScrollIntoView(el);
                ((IJavaScriptExecutor)_driver!).ExecuteScript("arguments[0].click();", el);
            }
            catch
            {
                // if even this fails, throw original behavior
                el.Click();
            }
        }

        private void ScrollIntoView(IWebElement el)
        {
            try
            {
                ((IJavaScriptExecutor)_driver!).ExecuteScript(
                    "arguments[0].scrollIntoView({block:'center', inline:'nearest'});", el);
            }
            catch { /* ignore */ }
        }

        private bool TryFindAllAcrossAllIframes(By by, out List<IWebElement> elements)
        {
            try
            {
                _driver!.SwitchTo().DefaultContent();
                var frames = _driver.FindElements(By.CssSelector("iframe"));

                foreach (var f in frames)
                {
                    if (!SafeDisplayed(f)) continue;

                    _driver.SwitchTo().Frame(f);

                    if (TryFindAllInCurrentContext(by, out elements, TimeSpan.FromMilliseconds(350)))
                        return true;

                    _driver.SwitchTo().DefaultContent();
                }
            }
            catch { /* ignore */ }

            elements = new List<IWebElement>();
            return false;
        }

        private static bool SafeDisplayed(IWebElement el)
        {
            try { return el.Displayed; }
            catch { return false; }
        }
        private static string? ExtractSelectValueFromTarget(List<string[]> selectors)
        {
            // Look for text=..., value=..., label=... inside Target list
            foreach (var g in selectors)
            {
                foreach (var s in g)
                {
                    var x = (s ?? "").Trim();
                    if (x.StartsWith("value=", StringComparison.OrdinalIgnoreCase)) return x;
                    if (x.StartsWith("label=", StringComparison.OrdinalIgnoreCase)) return x;
                    if (x.StartsWith("text=", StringComparison.OrdinalIgnoreCase)) return x.Substring("text=".Length).Trim();
                }
            }
            return null;
        }
        private string _cachedUrlForContext = "";
        private bool _cachedIsCrm = false;

        private bool IsCrmContext()
        {
            try
            {
                var url = _driver?.Url ?? "";

                // Cache by URL so we don't re-scan DOM every step
                if (string.Equals(url, _cachedUrlForContext, StringComparison.OrdinalIgnoreCase))
                    return _cachedIsCrm;

                bool isCrm = false;

                // Strongest check: CRM main page
                if (url.Contains("/main.aspx", StringComparison.OrdinalIgnoreCase) ||
                    url.Contains("main.aspx#", StringComparison.OrdinalIgnoreCase))
                {
                    isCrm = true;
                }

                // Backup check: CRM content iframe exists (only if driver is ready)
                if (!isCrm && _driver != null)
                {
                    var frames = _driver.FindElements(By.CssSelector("iframe[id*='contentIFrame'],iframe#AppLandingPage"));
                    if (frames.Count > 0) isCrm = true;
                }

                _cachedUrlForContext = url;
                _cachedIsCrm = isCrm;
                return isCrm;
            }
            catch
            {
                return false;
            }
        }

        private IWebElement LocatePreferFormControl(List<string[]> selectors)
        {
            // Prefer css/id/xpath and avoid aria label selectors first (they often locate labels not controls)
            var flat = selectors.SelectMany(g => g).Select(s => (s ?? "").Trim()).Where(s => s.Length > 0).ToList();

            var preferred = flat
                .Where(s =>
                    s.StartsWith("css=", StringComparison.OrdinalIgnoreCase) ||
                    s.StartsWith("id=", StringComparison.OrdinalIgnoreCase) ||
                    s.StartsWith("name=", StringComparison.OrdinalIgnoreCase) ||
                    s.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase) ||
                    s.StartsWith("//", StringComparison.Ordinal) ||
                    s.StartsWith("/", StringComparison.Ordinal))
                .ToList();

            // Fall back to all selectors if nothing preferred found
            var use = preferred.Count > 0 ? preferred : flat;

            return Locate(new List<string[]> { use.ToArray() });
        }

        private static (string cleanCommand, bool isOptional) ParseOptionalCommand(string rawCommand)
        {
            if (string.IsNullOrWhiteSpace(rawCommand))
                return ("", false);

            var cmd = rawCommand.Trim();

            //Optional ONLY if it starts with '*'
            bool isOptional = cmd.StartsWith("*");

            if (isOptional)
                cmd = cmd.Substring(1).Trim();

            //Safety: if someone typed trailing '*', strip it so command matches switch cases
            if (cmd.EndsWith("*"))
                cmd = cmd.TrimEnd('*').Trim();

            return (cmd.ToLowerInvariant(), isOptional);
        }

    }

}
