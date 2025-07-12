using System;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using System.Diagnostics;
using System.Collections.Generic;

// Alias to avoid ambiguity
using MailKitAuthEx = MailKit.Security.AuthenticationException;

class Program
{
    enum LogLevel { INFO, WARNING, ERROR }

    // Log file path initialized once per run
    static readonly string logFilePath = GetLogFilePath();

    static async Task Main(string[] args)
    {
        int totalUnsubscribed = 0;
        int totalDeleted = 0;
        int totalErrors = 0;

        try
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Welcome to InboxMaid! 🧹\n");
            Console.ResetColor();

            LogMessage("Application started", LogLevel.INFO);

            // Prompt for user credentials
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("Email address: ");
            Console.ResetColor();
            var email = Console.ReadLine();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("App password: ");
            Console.ResetColor();
            var password = ReadPassword();

            // Prompt for IMAP server (default to Yahoo)
            string defaultImapServer = "imap.mail.yahoo.com";
            Console.WriteLine("Press Enter to use the default server.");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"IMAP server (default: {defaultImapServer}): ");
            Console.ResetColor();
            var imapServer = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(imapServer))
                imapServer = defaultImapServer;

            int imapPort = 993; // Standard IMAP SSL port

            using var client = new ImapClient();

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nConnecting to IMAP server...");
            Console.ResetColor();

            try
            {
                await client.ConnectAsync(imapServer, imapPort, SecureSocketOptions.SslOnConnect);
                LogMessage($"Connected to IMAP server {imapServer}", LogLevel.INFO);
            }
            catch (SocketException)
            {
                ShowError("Unable to connect to the email server. Please check your internet connection and IMAP server address.");
                LogMessage("SocketException on ConnectAsync", LogLevel.ERROR);
                totalErrors++;
                return;
            }
            catch (Exception ex)
            {
                ShowError($"Unexpected error while connecting: {ex.Message}");
                LogMessage($"Exception on ConnectAsync: {ex}", LogLevel.ERROR);
                totalErrors++;
                return;
            }

            try
            {
                await client.AuthenticateAsync(email, password);
                LogMessage($"Authenticated user {email}", LogLevel.INFO);
            }
            catch (MailKitAuthEx)
            {
                ShowError("Authentication failed. Please verify your email and app password.");
                LogMessage("MailKit AuthenticationException on AuthenticateAsync", LogLevel.ERROR);
                totalErrors++;
                return;
            }
            catch (Exception ex)
            {
                ShowError($"Unexpected error during authentication: {ex.Message}");
                LogMessage($"Exception on AuthenticateAsync: {ex}", LogLevel.ERROR);
                totalErrors++;
                return;
            }

            var inbox = client.Inbox;
            await inbox.OpenAsync(FolderAccess.ReadWrite);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("Successfully connected! You have ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"{inbox.Unread}");
            Console.ResetColor();
            Console.Write(" unread emails out of ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write($"{inbox.Count}");
            Console.ResetColor();
            Console.WriteLine(" total emails in your inbox.");

            // Prompt for number of emails to scan
            int defaultEmailsToScan = 100;
            Console.WriteLine("Press Enter to use the default value.");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"How many recent emails should I scan? (default: {defaultEmailsToScan}): ");
            Console.ResetColor();
            var inputScanCount = Console.ReadLine();
            int emailsToScan;
            if (string.IsNullOrWhiteSpace(inputScanCount))
                emailsToScan = Math.Min(defaultEmailsToScan, inbox.Count);
            else if (int.TryParse(inputScanCount, out int userCount) && userCount > 0)
                emailsToScan = Math.Min(userCount, inbox.Count);
            else
                emailsToScan = Math.Min(defaultEmailsToScan, inbox.Count);

            // Prompt for mode
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("Choose mode: (1) Interactive (one-by-one) or (2) Batch (select multiple at once)? [choose 1 or 2, default is 1]: ");
            Console.ResetColor();
            var modeInput = Console.ReadLine()?.Trim();
            bool batchMode = (modeInput == "2");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nScanning your inbox for newsletters with unsubscribe links, please wait...");
            Console.ResetColor();

            if (!batchMode)
            {
                // Pre-scan newsletters only for interactive mode
                var uids = inbox.Search(SearchQuery.NotSeen).TakeLast(emailsToScan).ToList();

                var newsletters = new List<(UniqueId Uid, string From, string Subject, List<string> WebLinks, List<string> AllLinks)>();
                var seenNewsletters = new HashSet<string>();

                foreach (var uid in uids)
                {
                    var message = await inbox.GetMessageAsync(uid);
                    string uniqueKey = $"{message.Subject}|{message.From}";
                    if (seenNewsletters.Contains(uniqueKey))
                        continue;
                    seenNewsletters.Add(uniqueKey);

                    if (message.Headers.Contains("List-Unsubscribe"))
                    {
                        var unsubscribeHeader = message.Headers["List-Unsubscribe"];
                        var links = unsubscribeHeader
                            .Split(',')
                            .Select(h => h.Trim().Trim('<', '>'))
                            .ToList();

                        var webLinks = links
                            .Where(link => link.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (webLinks.Count > 0)
                            newsletters.Add((uid, message.From.ToString(), message.Subject, webLinks, links));
                    }
                }

                if (newsletters.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("🎉 No newsletters found in your recent emails! Your inbox is already clean.");
                    Console.ResetColor();
                    LogMessage("No newsletters found to unsubscribe.", LogLevel.INFO);
                }
                else
                {
                    (totalUnsubscribed, totalDeleted, totalErrors) = await RunInteractiveMode(newsletters, inbox, totalUnsubscribed, totalDeleted, totalErrors);
                }
            }
            else
            {
                (totalUnsubscribed, totalDeleted, totalErrors) = await RunBatchMode(inbox, emailsToScan, totalUnsubscribed, totalDeleted, totalErrors);
            }

            try
            {
                await inbox.ExpungeAsync();
            }
            catch (Exception ex)
            {
                LogMessage($"Exception during ExpungeAsync: {ex.Message}", LogLevel.WARNING);
                totalErrors++;
            }

            try
            {
                await client.DisconnectAsync(true);
            }
            catch (Exception ex)
            {
                LogMessage($"Exception during DisconnectAsync: {ex.Message}", LogLevel.WARNING);
                totalErrors++;
            }
        }
        catch (Exception ex)
        {
            ShowError($"An unexpected error occurred: {ex.Message}");
            LogMessage($"Unhandled exception in Main: {ex}", LogLevel.ERROR);
            totalErrors++;
        }

        // Write summary to log at end of run
        WriteSummary(totalUnsubscribed, totalDeleted, totalErrors);

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }

    static async Task<(int totalUnsubscribed, int totalDeleted, int totalErrors)> RunBatchMode(
        IMailFolder inbox,
        int batchSize,
        int totalUnsubscribed,
        int totalDeleted,
        int totalErrors)
    {
        bool batchModeActive = true;
        int currentOffset = 0;

        while (batchModeActive)
        {
            // Fetch unread emails starting at currentOffset  
            var uids = inbox.Search(SearchQuery.NotSeen)
                .Skip(currentOffset)
                .Take(batchSize)
                .ToList();

            // Build newsletters list for current batch
            var newsletters = new List<(UniqueId Uid, string From, string Subject, List<string> WebLinks, List<string> AllLinks)>();
            var seenNewsletters = new HashSet<string>();

            foreach (var uid in uids)
            {
                var message = await inbox.GetMessageAsync(uid);
                string uniqueKey = $"{message.Subject}|{message.From}";
                if (seenNewsletters.Contains(uniqueKey))
                    continue;
                seenNewsletters.Add(uniqueKey);

                if (message.Headers.Contains("List-Unsubscribe"))
                {
                    var unsubscribeHeader = message.Headers["List-Unsubscribe"];
                    var links = unsubscribeHeader
                        .Split(',')
                        .Select(h => h.Trim().Trim('<', '>'))
                        .ToList();

                    var webLinks = links
                        .Where(link => link.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (webLinks.Count > 0)
                        newsletters.Add((uid, message.From.ToString(), message.Subject, webLinks, links));
                }
            }

            if (newsletters.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\nNo newsletters found in this batch.");
                Console.ResetColor();
            }
            else
            {
                // Print newsletters with numbers in Cyan color
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\nNewsletters found:");
                for (int i = 0; i < newsletters.Count; i++)
                {
                    var n = newsletters[i];
                    Console.WriteLine($"{i + 1}) From: {n.From} | Subject: {n.Subject}");
                }
                Console.ResetColor();
            }

            Console.WriteLine("Choose an option:");
            Console.WriteLine("1) Unsubscribe from newsletters you select by their numbers"); // User picks specific newsletters
            Console.WriteLine("2) Unsubscribe from all newsletters");                      // Unsubscribe from all at once
            Console.WriteLine("3) Delete selected newsletters without unsubscribing");     // Delete emails without unsubscribing
            Console.WriteLine("4) Skip to next batch");                                    // Load next batch of emails
            Console.WriteLine("5) Restart scan");                                          // Restart scanning from beginning
            Console.WriteLine("6) Exit");                                                  // Exit batch unsubscribe mode
            Console.Write("Enter your choice (1-6): ");
            var choice = Console.ReadLine()?.Trim();

            switch (choice)
            {
                case "1":
                    Console.Write("Enter newsletter numbers separated by commas (no spaces needed), or type 'restart' to restart or 'exit' to exit batch unsubscribing: ");
                    var input = Console.ReadLine()?.Trim().ToLower();

                    if (input == "restart")
                    {
                        Console.WriteLine("Restarting scan...");
                        LogMessage("User requested scan restart in batch mode.", LogLevel.INFO);
                        currentOffset = 0; // reset offset to start over
                        break; // exit this case to restart loop
                    }
                    else if (input == "exit")
                    {
                        Console.WriteLine("Exiting batch unsubscribe mode.");
                        LogMessage("User exited batch unsubscribe mode.", LogLevel.INFO);
                        batchModeActive = false;
                        break; // exit this case and end loop
                    }
                    else
                    {
                        var selectedNumbers = input?.Split(',')
                            .Select(s => int.TryParse(s.Trim(), out int n) ? n : -1)
                            .Where(n => n > 0 && n <= newsletters.Count)
                            .Distinct()
                            .OrderByDescending(n => n) // Remove from highest index first
                            .ToList();

                        if (selectedNumbers != null && selectedNumbers.Count == 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("No valid newsletter numbers entered.");
                            Console.ResetColor();
                            break;
                        }

                        int batchUnsubscribed = 0;

                        if (selectedNumbers != null)
                            foreach (var num in selectedNumbers)
                            {
                                var selected = newsletters[num - 1];
                                var link = selected.WebLinks[0];
                                Console.WriteLine($"Opening: {link}");
                                try
                                {
                                    Process.Start(new ProcessStartInfo
                                    {
                                        FileName = link,
                                        UseShellExecute = true
                                    });
                                    await inbox.AddFlagsAsync(selected.Uid, MessageFlags.Deleted, true);
                                    totalUnsubscribed++;
                                    batchUnsubscribed++;
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("This newsletter email has been marked for deletion.");
                                    Console.ResetColor();
                                    LogMessage($"Unsubscribed from newsletter: {selected.Subject} from {selected.From}", LogLevel.INFO);
                                    newsletters.RemoveAt(num - 1);
                                }
                                catch (Exception ex)
                                {
                                    ShowError($"Failed to open browser: {ex.Message}");
                                    LogMessage($"Exception opening unsubscribe link: {ex}", LogLevel.ERROR);
                                    totalErrors++;
                                }
                            }

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"\nSuccessfully unsubscribed from {batchUnsubscribed} newsletter(s).");
                        Console.ResetColor();
                    }
                    break;

                case "2":
                    int allUnsubscribed = 0;
                    foreach (var newsletter in newsletters)
                    {
                        var link = newsletter.WebLinks[0];
                        Console.WriteLine($"Opening: {link}");
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = link,
                                UseShellExecute = true
                            });
                            await inbox.AddFlagsAsync(newsletter.Uid, MessageFlags.Deleted, true);
                            totalUnsubscribed++;
                            allUnsubscribed++;
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("This newsletter email has been marked for deletion.");
                            Console.ResetColor();
                            LogMessage($"Unsubscribed from newsletter: {newsletter.Subject} from {newsletter.From}", LogLevel.INFO);
                        }
                        catch (Exception ex)
                        {
                            ShowError($"Failed to open browser: {ex.Message}");
                            LogMessage($"Exception opening unsubscribe link: {ex}", LogLevel.ERROR);
                            totalErrors++;
                        }
                    }
                    newsletters.Clear();

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"\nSuccessfully unsubscribed from all {allUnsubscribed} newsletters.");
                    Console.ResetColor();
                    break;

                case "3":
                    Console.Write("Enter newsletter numbers to delete, separated by commas (no spaces needed), or type 'restart' to restart or 'exit' to exit batch unsubscribing: ");
                    var deleteInput = Console.ReadLine()?.Trim().ToLower();

                    if (deleteInput == "restart")
                    {
                        Console.WriteLine("Restarting scan...");
                        LogMessage("User requested scan restart in batch mode.", LogLevel.INFO);
                        currentOffset = 0;
                        break;
                    }
                    else if (deleteInput == "exit")
                    {
                        Console.WriteLine("Exiting batch unsubscribe mode.");
                        LogMessage("User exited batch unsubscribe mode.", LogLevel.INFO);
                        batchModeActive = false;
                        break;
                    }
                    else
                    {
                        var deleteNumbers = deleteInput?.Split(',')
                            .Select(s => int.TryParse(s.Trim(), out int n) ? n : -1)
                            .Where(n => n > 0 && n <= newsletters.Count)
                            .Distinct()
                            .OrderByDescending(n => n)
                            .ToList();

                        if (deleteNumbers != null && deleteNumbers.Count == 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("No valid newsletter numbers entered.");
                            Console.ResetColor();
                            break;
                        }

                        int batchDeleted = 0;

                        if (deleteNumbers != null)
                            foreach (var num in deleteNumbers)
                            {
                                var selected = newsletters[num - 1];
                                Console.WriteLine($"Deleting email from: {selected.From} | Subject: {selected.Subject}");
                                try
                                {
                                    await inbox.AddFlagsAsync(selected.Uid, MessageFlags.Deleted, true);
                                    totalDeleted++;
                                    batchDeleted++;
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("This newsletter email has been marked for deletion.");
                                    Console.ResetColor();
                                    LogMessage($"Deleted newsletter without unsubscribing: {selected.Subject} from {selected.From}", LogLevel.INFO);
                                    newsletters.RemoveAt(num - 1);
                                }
                                catch (Exception ex)
                                {
                                    ShowError($"Failed to delete email: {ex.Message}");
                                    LogMessage($"Exception deleting email: {ex}", LogLevel.ERROR);
                                    totalErrors++;
                                }
                            }

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"\nSuccessfully deleted {batchDeleted} newsletter(s) without unsubscribing.");
                        Console.ResetColor();
                    }
                    break;

                case "4":
                    currentOffset += batchSize;
                    if (currentOffset >= inbox.Count)
                    {
                        Console.WriteLine("Reached the end of your inbox. No more batches.");
                        currentOffset -= batchSize; // reset offset to last batch    
                    }
                    else
                    {
                        Console.WriteLine("Loading next batch of newsletters...");
                    }
                    break;

                case "5":
                    Console.WriteLine("Restarting scan...");
                    LogMessage("User requested scan restart in batch mode.", LogLevel.INFO);
                    currentOffset = 0;
                    break;

                case "6":
                    Console.WriteLine("Exiting batch unsubscribe mode.");
                    LogMessage("User exited batch unsubscribe mode.", LogLevel.INFO);
                    batchModeActive = false;
                    break;

                default:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid choice. Please enter a number between 1 and 6.");
                    Console.ResetColor();
                    break;
            }

            Console.WriteLine($"\nYou unsubscribed from {totalUnsubscribed} newsletter(s) and deleted {totalDeleted} newsletter(s) without unsubscribing in total.");
        }

        return (totalUnsubscribed, totalDeleted, totalErrors);
    }

    static async Task<(int totalUnsubscribed, int totalDeleted, int totalErrors)> RunInteractiveMode(
        List<(UniqueId Uid, string From, string Subject, List<string> WebLinks, List<string> AllLinks)> newsletters,
        IMailFolder inbox,
        int totalUnsubscribed,
        int totalDeleted,
        int totalErrors)
    {
        foreach (var newsletter in newsletters)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine("📧 Newsletter:");
            Console.WriteLine($"   From: {newsletter.From}");
            Console.WriteLine($"   Subject: {newsletter.Subject}");
            Console.WriteLine("   Unsubscribe options:");
            foreach (var link in newsletter.AllLinks)
            {
                if (link.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    Console.WriteLine($"      [Web]    {link}");
                else if (link.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                    Console.WriteLine($"      [Email]  {link}");
                else
                    Console.WriteLine($"      [Other]  {link}");
            }
            Console.WriteLine();
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("Open unsubscribe link (y), delete email (d), skip (n), or exit: ");
            Console.ResetColor();
            var response = Console.ReadLine()?.Trim().ToLower();

            if (response == "y")
            {
                var link = newsletter.WebLinks[0];
                Console.WriteLine($"Opening: {link}");
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = link,
                        UseShellExecute = true
                    });
                    totalUnsubscribed++;
                    await inbox.AddFlagsAsync(newsletter.Uid, MessageFlags.Deleted, true);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("This newsletter email has been marked for deletion.");
                    Console.ResetColor();
                    LogMessage($"Unsubscribed from newsletter: {newsletter.Subject} from {newsletter.From}", LogLevel.INFO);
                }
                catch (Exception ex)
                {
                    ShowError($"Failed to open browser: {ex.Message}");
                    LogMessage($"Exception opening unsubscribe link: {ex}", LogLevel.ERROR);
                    totalErrors++;
                }
            }
            else if (response == "d")
            {
                await inbox.AddFlagsAsync(newsletter.Uid, MessageFlags.Deleted, true);
                totalDeleted++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("This newsletter email has been marked for deletion (without unsubscribing).");
                Console.ResetColor();
                LogMessage($"Deleted newsletter without unsubscribing: {newsletter.Subject} from {newsletter.From}", LogLevel.INFO);
            }
            else if (response == "exit")
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Exiting newsletter review.");
                Console.ResetColor();
                LogMessage("User exited interactive unsubscribe mode.", LogLevel.INFO);
                break;
            }
            // 'n' or any other input skips the newsletter silently
        }

        Console.WriteLine($"\nCongratulations! InboxMaid has helped you unsubscribe from {totalUnsubscribed} newsletter{(totalUnsubscribed == 1 ? "" : "s")} and delete {totalDeleted} newsletter{(totalDeleted == 1 ? "" : "s")} without unsubscribing. Welcome to a cleaner Inbox!");

        return (totalUnsubscribed, totalDeleted, totalErrors);
    }

    static void ShowError(string message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    static void LogMessage(string message, LogLevel level = LogLevel.INFO)
    {
        try
        {
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}\n";
            File.AppendAllText(logFilePath, logEntry);
        }
        catch
        {
            // Ignore logging errors to avoid crashing the app
        }
    }

    static void WriteSummary(int totalUnsubscribed, int totalDeleted, int totalErrors)
    {
        try
        {
            string summary = $"\n=== Session Summary ===\n" +
                             $"Total unsubscribed: {totalUnsubscribed}\n" +
                             $"Total deleted without unsubscribing: {totalDeleted}\n" +
                             $"Total errors: {totalErrors}\n" +
                             $"Session ended at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                             $"======================\n";

            File.AppendAllText(logFilePath, summary);
        }
        catch
        {
            // Ignore logging errors
        }
    }

    static string GetLogFilePath()
    {
        string logDir = AppDomain.CurrentDomain.BaseDirectory;
        string datePart = DateTime.Now.ToString("yyyyMMdd");
        string baseFileName = $"error_{datePart}.log";
        string fullPath = Path.Combine(logDir, baseFileName);

        if (!File.Exists(fullPath))
        {
            return fullPath;
        }
        else
        {
            // If file exists, add timestamp to avoid overwrite
            string timePart = DateTime.Now.ToString("HHmmss");
            string timestampedFileName = $"error_{datePart}_{timePart}.log";
            return Path.Combine(logDir, timestampedFileName);
        }
    }

    // Helper method to read password without echoing
    static string ReadPassword()
    {
        string password = "";
        ConsoleKeyInfo info;
        do
        {
            info = Console.ReadKey(true);
            if (info.Key != ConsoleKey.Enter && info.Key != ConsoleKey.Backspace)
            {
                password += info.KeyChar;
                Console.Write("*");
            }
            else if (info.Key == ConsoleKey.Backspace && password.Length > 0)
            {
                password = password.Substring(0, password.Length - 1);
                Console.Write("\b \b");
            }
        } while (info.Key != ConsoleKey.Enter);
        Console.WriteLine();
        return password;
    }
}