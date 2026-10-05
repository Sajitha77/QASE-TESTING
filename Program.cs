using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using TestAutomationFramework.Framework;

namespace TestAutomationFramework
{
    public class Program
    {
        
        static int Main(string[] args)
        {
            
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                PrintStartupInfo();

                // If command-line args are provided, go straight into parameter mode.
                // This makes pipeline execution non-interactive.
                if (args != null && args.Length > 0)
                {
                    return RunWithParameters(args);
                }
                
                while (true)
                {
                    Console.WriteLine("Select execution mode:");
                    Console.WriteLine("1. CLI Menu");
                    Console.WriteLine("2. Debug Mode");
                    Console.WriteLine("3. Parameter-Based");
                    Console.WriteLine("4. GUI (Coming Soon)");
                    Console.WriteLine("5. Exit");
                    Console.Write("> ");

                    string? mode = Console.ReadLine()?.Trim();

                    switch (mode)
                    {
                        case "1":
                            RunWithCLIMenu();
                            break;
                    
                        case "2":
                            DebugRunner.RunDebugMenu();
                            break;

                        case "3":
                            Console.WriteLine("Parameter-based execution requires command-line arguments.");
                            Console.WriteLine("Example:");
                            Console.WriteLine(@"TestAutomationFramework.exe --project ""DCP - PPD"" --test ""MyTest.xlsx"" --headless true --incognito false");
                            Console.WriteLine(@"Or:");
                            Console.WriteLine(@"TestAutomationFramework.exe --project ""DCP - PPD"" --runAll true --headless true --incognito false");
                            break;

                        case "4":
                            Console.WriteLine("GUI mode is currently disabled. Please select a different option");
                            break;
                        case "5":
                            Console.WriteLine("Exiting program...");
                            return 0;
                        default:
                            Console.WriteLine("Invalid selection. Please select from the above options");
                            Console.WriteLine();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected startup error: {ex.Message}");
                return 1;
            }
        }

        static void PrintStartupInfo()
        {
            Console.WriteLine($"{AppInfo.ProductName}");
            Console.WriteLine($"Version : {AppInfo.Version}");
            Console.WriteLine($"Built On: {AppInfo.BuiltOn}");
            Console.WriteLine();
        }

        static void RunWithCLIMenu()
        {
            
           
            bool headless = false;
            bool incognito = false;  
            bool windowsAuth = false;
            string? authServerAllowlist = null;
            while(true){
            windowsAuth = false;
            authServerAllowlist = null;
            Console.WriteLine("Select browser mode.");
            Console.WriteLine("1. Incognito browser");
            Console.WriteLine("2. Headless + Incognito (No UI)");
            Console.WriteLine("3. Normal browser");
            Console.WriteLine("4. Headless browser (No UI)");
            Console.WriteLine("5. Windows authentication (current Windows user)");
            Console.WriteLine("6. Back");
            Console.Write("> ");
            string? headlessIncogInput = Console.ReadLine()?.Trim();
            switch (headlessIncogInput)
            {
            case "1":
                headless = false;
                incognito = true;
                break;

            case "2":
                headless = true;
                incognito = true;
                break;
            
            case "3":
                headless = false;
                incognito = false;
                break;
            case "4":
                headless = true;
                incognito = false;
                break;
            case "5":
                headless = false;
                incognito = false;
                windowsAuth = true;
                if (!OperatingSystem.IsWindows())
                {
                    Console.WriteLine("Windows authentication mode is available on Windows only.");
                    continue;
                }
                Console.WriteLine("Chrome will use the Windows account running QASE.");
                Console.WriteLine("Enter the trusted site hostname(s), separated by commas.");
                Console.WriteLine("Example: dlls-slrd-staging.forces.mil.ca");
                Console.Write("Hostnames (blank to go back): ");
                string? hostsInput = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(hostsInput))
                    continue;
                try
                {
                    authServerAllowlist = BrowserOptionsFactory.ValidateAuthentication(true, false, hostsInput);
                }
                catch (ArgumentException ex)
                {
                    Console.WriteLine(ex.Message);
                    continue;
                }
                break;
            case "6":
                return;
            default:
                Console.WriteLine("Invalid selection. Please try again.");
                continue;
            }
                
                    
            if (!TryResolveTestingCentrePath(out string testingCentrePath, out string[] searchedPaths))
            {
                PrintTestingCentreNotFound(searchedPaths);
                Console.WriteLine("\nPress Enter to exit...");
                Console.ReadLine();
                return;
            }

            while (true)
            {
                var projects = Directory.GetDirectories(testingCentrePath);

                if (projects.Length == 0)
                {
                    Console.WriteLine("No projects found in TestingCentre.");
                    return;
                }

                Console.WriteLine("\nSelect a project to test:");
                for (int i = 0; i < projects.Length; i++)
                {
                    Console.WriteLine($"{i + 1}. {Path.GetFileName(projects[i])}");
                }

                Console.WriteLine($"{projects.Length + 1}. Back");
                Console.WriteLine($"{projects.Length + 2}. Main Menu");
                Console.Write("> ");

                if (!int.TryParse(Console.ReadLine(), out int projectIndex) ||
                projectIndex < 1 || projectIndex > projects.Length + 2)
                {
                    Console.WriteLine("Invalid selection.");
                    continue;
                }

                if (projectIndex == projects.Length + 1)
                {
                    Console.WriteLine("Returning to previous menu...");
                    break;
                }
                if (projectIndex == projects.Length + 2)
                {
                    Console.WriteLine("Returning to main menu...");
                    return;
                }
                string selectedProjectPath = projects[projectIndex - 1];
                string inputPath = Path.Combine(selectedProjectPath, "Input");

                

                while (true)
                {
                    if (!Directory.Exists(inputPath))
                {
                    Console.WriteLine($"\nInput folder not found for project: {Path.GetFileName(selectedProjectPath)}");
                    Console.WriteLine("Press Enter to return to project selection...");
                    Console.ReadLine();
                    break;
                }

                var initialJsonFiles = Directory.GetFiles(inputPath, "*.json", SearchOption.TopDirectoryOnly);
                Console.WriteLine($"Found {initialJsonFiles.Length} JSON file(s) in '{inputPath}'");

                TryConvertJsonFiles(inputPath, Path.Combine(selectedProjectPath, "Output"));

                var testFiles = GetRunnableTestFiles(inputPath);
                var directories = Getdirectories(inputPath);

                if (testFiles.Count == 0 && directories.Count == 0)
                {
                    Console.WriteLine("\nNo test cases or directories found in Input folder.");
                    Console.WriteLine("Press Enter to return to project selection...");
                    Console.ReadLine();
                    break;
                }


                    Console.WriteLine("\nSelect a test case to run, or a directory to enter:");
                    Console.WriteLine("1. Run all test cases");
                    for (int i=0; i<directories.Count; i++)
                        {
                            Console.WriteLine($"{i+2}. {Path.GetFileName(directories[i])} (dir)");
                        }

                    for (int i = 0; i < testFiles.Count; i++)
                    {
                        Console.WriteLine($"{i + 2 + directories.Count}. {Path.GetFileName(testFiles[i])}");
                    }

                    Console.WriteLine($"{testFiles.Count + directories.Count + 2}. Back");
                    Console.WriteLine($"{testFiles.Count + directories.Count + 3}. Main Menu");
                    Console.Write("> ");

                    string input = Console.ReadLine()?.Trim() ?? "";

                    if (!int.TryParse(input, out int selection) ||
                    selection < 1 || selection > testFiles.Count + directories.Count + 3)
                    {
                        Console.WriteLine("Invalid selection. Try again.");
                        continue;
                    }
                    if (selection == testFiles.Count + directories.Count + 2)
                    {
                        break;
                    }
                    if (selection == testFiles.Count + directories.Count+ 3)
                    {
                        return;
                    }
                    

                   if (selection>1 && selection < directories.Count + 2)
                        {
                            inputPath = Path.Combine(inputPath, directories[selection-2]);
                        }
                    else{
                    string? selectedTestFile = selection == 1 ? null : testFiles[selection - directories.Count - 2];
                    bool runAll = selection == 1;

                    while (true)
                    {
                        int exitCode = ExecuteProjectTests(
    
     selectedProjectPath,
     selectedTestFile,
     runAll,
     headless,
     incognito,
     0,
    
     skipJsonConversion: true,
     pauseAtEnd: true,
     inputPath,
     windowsAuth,
     authServerAllowlist
 );

                        // For menu mode, show invalid configuration issues but return to menu instead of quitting.
                        if (exitCode == 2)
                        {
                            Console.WriteLine("Execution ended with invalid configuration or input.");
                        }

                        Console.WriteLine("\nWhat would you like to do next?");
                        Console.WriteLine("1. Rerun same test(s)");
                        Console.WriteLine("2. Select another test");
                        Console.WriteLine("3. Main Menu");
                        Console.Write("> ");

                        string? nextAction = Console.ReadLine()?.Trim();

                        if (nextAction == "1")
                        {
                            continue;
                        }
                        else if (nextAction == "2")
                        {
                            break;
                        }
                        else if (nextAction == "3")
                        {
                            Console.WriteLine("Returning to Main menu...");
                            return;
                        }
                        else
                        {
                            Console.WriteLine("Invalid selection. Returning to test selection...");
                            break;
                        }
                    }
                    }

                }
            }
        }
        }

        static int RunWithParameters(string[] args)
        {
            String? inputPath=null;
            var parsedArgs = ParseArguments(args);

            if (parsedArgs.ContainsKey("help") || parsedArgs.ContainsKey("?"))
            {
                PrintParameterUsage();
                return 0;
            }

            if (!parsedArgs.TryGetValue("project", out string? projectName) || string.IsNullOrWhiteSpace(projectName))
            {
                Console.WriteLine("ERROR: Missing required parameter '--project'.");
                PrintParameterUsage();
                return 2;
            }

            bool runAll = GetBoolArg(parsedArgs, "runAll", false);
            bool headless = GetBoolArg(parsedArgs, "headless", false);
            int stepDelayMs = GetIntArg(parsedArgs, "stepDelayMs", 0);
            bool incognito = GetBoolArg(parsedArgs, "incognito", false);
            bool windowsAuth = GetBoolArg(parsedArgs, "windowsAuth", false);
            parsedArgs.TryGetValue("authServerAllowlist", out string? authServerAllowlist);
            try
            {
                authServerAllowlist = BrowserOptionsFactory.ValidateAuthentication(windowsAuth, incognito, authServerAllowlist);
                BrowserOptionsFactory.EnsureWindowsAuthenticationSupported(windowsAuth);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is PlatformNotSupportedException)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                return 2;
            }
            parsedArgs.TryGetValue("test", out string? testFileName);

            if (!runAll && string.IsNullOrWhiteSpace(testFileName))
            {
                Console.WriteLine("ERROR: You must provide either '--test \"fileName.xlsx\"' or '--runAll true'.");
                PrintParameterUsage();
                return 2;
            }

            if (runAll && !string.IsNullOrWhiteSpace(testFileName))
            {
                Console.WriteLine("ERROR: Use either '--test' or '--runAll true', not both.");
                PrintParameterUsage();
                return 2;
            }

            if (!TryResolveTestingCentrePath(out string testingCentrePath, out string[] searchedPaths))
            {
                PrintTestingCentreNotFound(searchedPaths);
                return 2;
            }

            string? selectedProjectPath = Directory
            .GetDirectories(testingCentrePath)
            .FirstOrDefault(dir =>
            string.Equals(Path.GetFileName(dir), projectName, StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(selectedProjectPath))
            {
                Console.WriteLine($"ERROR: Project '{projectName}' was not found in TestingCentre.");
                Console.WriteLine("Available projects:");

                foreach (var dir in Directory.GetDirectories(testingCentrePath))
                {
                    Console.WriteLine($"- {Path.GetFileName(dir)}");
                }

                return 2;
            }

            return ExecuteProjectTests(
                

    selectedProjectPath,

    testFileName,

    runAll,

    headless,

    incognito,

    stepDelayMs,

    skipJsonConversion: false,

    pauseAtEnd: false,
    inputPath,
    windowsAuth,
    authServerAllowlist

);

        }

        static bool TryResolveTestingCentrePath(out string path, out string[] searchedPaths)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            searchedPaths = new[]
            {
                Path.Combine(baseDir, "TestingCentre"),
                Path.Combine(Directory.GetCurrentDirectory(), "TestingCentre"),
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "TestingCentre")),
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "TestingCentre"))
            };

            path = searchedPaths.FirstOrDefault(Directory.Exists) ?? string.Empty;
            return !string.IsNullOrWhiteSpace(path);
        }

        static void PrintTestingCentreNotFound(string[] searchedPaths)
        {
            Console.WriteLine("ERROR: TestingCentre folder cannot be found.");
            Console.WriteLine("Searched locations:");
            foreach (var candidate in searchedPaths)
            {
                Console.WriteLine($"- {candidate}");
            }
        }

        static int ExecuteProjectTests(

           string selectedProjectPath,
           string? selectedTestFileName,
           bool runAll,
           bool headless,
           bool incognito,
           int stepDelayMs,
           bool skipJsonConversion,
           bool pauseAtEnd,
           string? inputPath,
           bool windowsAuth = false,
           string? authServerAllowlist = null)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputRoot = Path.Combine(selectedProjectPath, "Input");
            // Parameter mode has no interactive folder selection.
            inputPath ??= inputRoot;
            string outputPath = Path.Combine(selectedProjectPath, "Output");
            string logsPath = Path.Combine(baseDir, "developer-logs");
            string? requestedTestFileName = string.IsNullOrWhiteSpace(selectedTestFileName)
                ? null
                : Path.GetFileName(selectedTestFileName.Trim());

            if (!Directory.Exists(inputRoot))
            {
                Console.WriteLine($"ERROR: Input folder not found at: {inputPath}");
                return 2;
            }

            Directory.CreateDirectory(outputPath);
            Directory.CreateDirectory(logsPath);

            string projectName = Path.GetFileName(selectedProjectPath);
            string devLogPath = Path.Combine(logsPath, $"{projectName}_log.txt");

            using var logWriter = new StreamWriter(devLogPath, append: true);

            void Log(string message)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                logWriter.WriteLine($"[{timestamp}] {message}");
                logWriter.Flush();
            }

            Log("=== New Test Session Started ===");
            Log($"Selected Project: {projectName}");
            Log($"Execution Mode: {(runAll ? "Run All" : "Single Test")}");
            Log($"Headless: {headless}");
            Log($"Incognito: {incognito}");
            Log($"Windows authentication: {windowsAuth}");
            if (windowsAuth)
            {
                Log($"Authentication server allowlist: {authServerAllowlist}");
                Console.WriteLine("Windows authentication: using the Windows account running QASE.");
                Console.WriteLine($"Trusted authentication hosts: {authServerAllowlist}");
            }
            if (!string.IsNullOrWhiteSpace(requestedTestFileName))
            {
                Log($"Requested Test File: {requestedTestFileName}");
            }

            if (!skipJsonConversion)
            {
                var initialJsonFiles = Directory.GetFiles(inputPath, "*.json", SearchOption.TopDirectoryOnly);
                Console.WriteLine($"Found {initialJsonFiles.Length} JSON file(s) in '{inputPath}'");

                try
                {
                    JsonToExcelConverter.ConvertAll(inputPath, outputPath);
                    Log("Converted any JSON files to Excel.");
                }
                catch (Exception ex)
                {
                    Log($"ERROR during JSON conversion: {ex.Message}\n{ex.StackTrace}");
                }
            }
            List<string>? testFiles = null;
            if (runAll){
                        var directories = Directory.GetFileSystemEntries(inputPath);
                        testFiles = getAllTestFiles(inputPath);
                        }
            else{
                var directories = Directory.GetFileSystemEntries(inputPath);
                     testFiles = GetRunnableTestFiles(inputPath);
                } 

            if (testFiles.Count == 0)
            {
                Console.WriteLine("ERROR: No test cases found in Input folder.");
                Log("No test cases found in Input folder.");
                return 2;
            }

            int passed = 0;
            int failed = 0;
            List<string> failedTests = new();
            List<string> passedTests = new();

            if (runAll)
            {   
                testFiles=getAllTestFiles(inputPath);
                foreach (var file in testFiles)
                {
                    RunSingleTest(file, headless, incognito, stepDelayMs, ref passed, ref failed, ref failedTests, ref passedTests, Log, windowsAuth, authServerAllowlist);
                }
            }
            else
            {
                string? chosenFile = testFiles.FirstOrDefault(f =>
                string.Equals(Path.GetFileName(f), requestedTestFileName, StringComparison.OrdinalIgnoreCase));

                if (string.IsNullOrWhiteSpace(chosenFile))
                {
                    Console.WriteLine($"ERROR: Test file '{requestedTestFileName}' was not found in project '{projectName}'.");
                    Console.WriteLine("Available test files:");

                    foreach (var file in testFiles)
                    {
                        Console.WriteLine($"- {Path.GetFileName(file)}");
                    }

                    Log($"Requested test file not found: {requestedTestFileName}");
                    return 2;
                }

                RunSingleTest(chosenFile, headless, incognito, stepDelayMs, ref passed, ref failed, ref failedTests, ref passedTests, Log, windowsAuth, authServerAllowlist);
            }
            Console.WriteLine($"\nTest Summary: {passed} passed, {failed} failed.");
            if (runAll)
            {
                var failedTestNames = failedTests.Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
                var passedTestNames = passedTests.Where(t => !string.IsNullOrWhiteSpace(t)).ToList();

                Console.WriteLine("Tests failed:");
                if (failedTestNames.Count > 0)
                {
                    foreach (var testName in failedTestNames)
                    {
                        Console.WriteLine($" - {testName}");
                    }
                }
                else
                {
                    Console.WriteLine(" none");
                }

                Console.WriteLine("Tests passed:");
                if (passedTestNames.Count > 0)
                {
                    foreach (var testName in passedTestNames)
                    {
                        Console.WriteLine($" - {testName}");
                    }
                }
                else
                {
                    Console.WriteLine(" none");
                }
            }
            Console.WriteLine($"Test results saved to: {Path.GetFullPath(outputPath)}");
            Console.WriteLine($"Developer log saved to: {Path.GetFullPath(devLogPath)}");

            Log($"Test Summary: {passed} passed, {failed} failed.");
            if (runAll)
            {
                Log($"Tests failed: {(failed > 0 ? string.Join(", ", failedTests.Where(t => !string.IsNullOrWhiteSpace(t))) : "none")}");
                Log($"Tests passed: {(passed > 0 ? string.Join(", ", passedTests.Where(t => !string.IsNullOrWhiteSpace(t))) : "none")}");
            }
            Log($"Test results saved to: {Path.GetFullPath(outputPath)}");
            Log($"Developer log saved to: {Path.GetFullPath(devLogPath)}");

            if (pauseAtEnd)
            {
                Console.WriteLine("\nPress Enter to continue");
                Console.ReadLine();
            }

            return failed > 0 ? 1 : 0;
        }

        static List<string> GetRunnableTestFiles(string inputPath)
        {
            if (!Directory.Exists(inputPath))
            {
                return new List<string>();
            }

            return Directory.GetFiles(inputPath)
            .Where(f =>
            f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".puppateer", StringComparison.OrdinalIgnoreCase))
            .OrderBy(Path.GetFileName)
            .ToList();
        }

        public static List<string> Getdirectories(string inputPath)
        {
           if (!Directory.Exists(inputPath))
            {
                return new List<string>();
            }
            return Directory.GetFileSystemEntries(inputPath)
            .Where(f =>
            Directory.Exists(f))
            .OrderBy(Path.GetFileName)
            .ToList();
        }

        static void TryConvertJsonFiles(string inputPath, string outputPath)
        {
            try
            {
                JsonToExcelConverter.ConvertAll(inputPath, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR during JSON conversion: {ex.Message}");
            }
        }

        static Dictionary<string, string> ParseArguments(string[] args)
        {
            var parsed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length; i++)
            {
                string current = args[i];

                if (!current.StartsWith("--"))
                {
                    continue;
                }

                string key = current.Substring(2);

                // Support flags like: --help
                if (i == args.Length - 1 || args[i + 1].StartsWith("--"))
                {
                    parsed[key] = "true";
                    continue;
                }

                parsed[key] = args[i + 1];
                i++;
            }

            return parsed;
        }

        static bool GetBoolArg(Dictionary<string, string> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out string? value) || string.IsNullOrWhiteSpace(value))
            {
                return defaultValue;
            }

            return value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }

        static void PrintParameterUsage()
        {
            Console.WriteLine("\nParameter-Based Usage:");
            Console.WriteLine(@"Single test:");
            Console.WriteLine(@" TestAutomationFramework.exe --project ""DCP - PPD"" --test ""14525_ Approval Checklist functionality - Unit Rep (1).xlsx"" --headless true --incognito false");
            Console.WriteLine();
            Console.WriteLine(@"Run all tests:");
            Console.WriteLine(@" TestAutomationFramework.exe --project ""DCP - PPD"" --runAll true --headless true --incognito false");
            Console.WriteLine();
            Console.WriteLine("Supported parameters:");
            Console.WriteLine(" --project Required. Project folder name inside TestingCentre.");
            Console.WriteLine(" --test Optional. Exact file name of one test case to run.");
            Console.WriteLine(" --runAll Optional. true/false. Runs all test cases in the project.");
            Console.WriteLine(" --headless Optional. true/false.");
            Console.WriteLine(" --incognito Optional. true/false.");
            Console.WriteLine(" --windowsAuth Optional. true/false. Use the current Windows account; incompatible with Incognito.");
            Console.WriteLine(" --authServerAllowlist Required with --windowsAuth true. Comma-separated exact trusted hostnames.");
            Console.WriteLine(@" Windows sign-in example: TestAutomationFramework.exe --project ""DLLS"" --test ""DLLS_Initial_Test_Case.xlsx"" --windowsAuth true --authServerAllowlist dlls-slrd-staging.forces.mil.ca");
            Console.WriteLine(" --help Show usage.");
            Console.WriteLine();
        }

        static void RunSingleTest(string file, bool headless, bool incognito, int stepDelayMs, ref int passed, ref int failed, ref List<string> failedTests, ref List<string> passedTests, Action<string> Log,
            bool windowsAuth = false, string? authServerAllowlist = null)
        {
            string fileName = Path.GetFileName(file);
            Log($"Running script: {fileName}");

            try
            {
                var runner = new TestSuiteRunner();
                var result = runner.Run(file, headless, incognito, stepDelayMs, windowsAuth, authServerAllowlist);

                if (result.Passed)
                {
                    passed++;
                    passedTests.Add(result.FileName);
                    Log($"PASSED: {result.FileName}");
                }
                else
                {
                    failed++;
                    failedTests.Add(result.FileName);
                    Log($"FAILED: {result.FileName}");

                    foreach (var err in result.Errors)
                    {
                        Log(" - " + err);
                    }

                    if (!string.IsNullOrEmpty(result.ScreenshotPath))
                    {
                        Log($"Screenshot: {result.ScreenshotPath}");
                    }
                }

                Log($"Output Excel: {result.OutputFilePath}");
            }
            catch (Exception ex)
            {
                Log($"ERROR running {fileName}: {ex.Message}\n{ex.StackTrace}");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed: {fileName}");
                if (ex.Message.Contains("unable to locate selenium manager binary", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Unable to locate the Chrome driver.");
                }
                else
                {
                    Console.WriteLine(ex.Message);
                }
                if (ex.Message.Contains("only supports Chrome version", StringComparison.OrdinalIgnoreCase) ||
                    ex.Message.Contains("session not created", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("ChromeDriver version mismatch detected. Update/rebuild with a ChromeDriver version matching your installed Chrome.");
                }
                else if (ex.Message.Contains("cannot find Chrome binary", StringComparison.OrdinalIgnoreCase) ||
                         ex.Message.Contains("Chrome failed to start", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Chrome is not installed or could not be found. Please install Google Chrome and try again.");
                }
                else if (ex.Message.Contains("chromedriver", StringComparison.OrdinalIgnoreCase) &&
                         (ex.Message.Contains("not found", StringComparison.OrdinalIgnoreCase) ||
                          ex.Message.Contains("does not exist", StringComparison.OrdinalIgnoreCase) ||
                          ex.Message.Contains("needs to be in PATH", StringComparison.OrdinalIgnoreCase)))
                {
                    Console.WriteLine("ChromeDriver could not be located. Place chromedriver.exe in the same folder as this application, or ensure internet access for automatic download.");
                }
                else if (ex.Message.Contains("download", StringComparison.OrdinalIgnoreCase) ||
                         ex.Message.Contains("http", StringComparison.OrdinalIgnoreCase) ||
                         ex.Message.Contains("network", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Could not auto-download ChromeDriver. Check your internet connection or place a matching chromedriver.exe next to this application.");
                }
                Console.ResetColor();
                failed++;
                failedTests.Add(fileName);
            }
        }
        static List<string> getAllTestFiles(string inputPath){
            if (!Directory.Exists(inputPath)){
                return new List<string>();
                }
            List<string> files = new List<string>();
            collectAllTestFiles(inputPath, files);
            return files;
        }
        static void collectAllTestFiles(string subdirectory, List<string> files)
        {    if (!Directory.Exists(subdirectory)){
                return;
                }
            
            files.AddRange(GetRunnableTestFiles(subdirectory));
            foreach (var dir in Getdirectories(subdirectory)){
                collectAllTestFiles(dir, files);
            }
            return;
        }
        static int GetIntArg(Dictionary<string, string> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out string? value) || string.IsNullOrWhiteSpace(value))
            {
                return defaultValue;
            }
            return int.TryParse(value, out int parsed) ? parsed : defaultValue;
        }


    }
}
