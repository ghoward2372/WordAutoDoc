using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;

namespace DocumentProcessor.Services
{
    public class GrammarChecker
    {
        private readonly Application? _wordApp;
        private readonly Document? _doc;
        private readonly bool _isComSupported;

        public GrammarChecker(string filePath)
        {
            try
            {
                // Check if running on Windows and COM is supported
                _isComSupported = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
                if (!_isComSupported)
                {
                    Console.WriteLine("Warning: Grammar checking is only supported on Windows platforms.");
                    return;
                }

                _wordApp = new Application { Visible = false };
                _doc = _wordApp.Documents.Open(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Grammar checking is not available - {ex.Message}");
                _isComSupported = false;

                // Ensure cleanup in case of partial initialization
                if (_doc != null) _doc.Close();
                if (_wordApp != null) _wordApp.Quit();
            }
        }

        public void CheckAndFixGrammar()
        {
            if (!_isComSupported || _doc == null)
            {
                Console.WriteLine("Grammar checking skipped - not supported on this platform.");
                return;
            }

            try
            {
                Console.WriteLine("\n=== Grammar Check Started ===");

                // First check - List all errors
                Console.WriteLine("\nInitial Grammar Errors:");
                ListGrammarErrors();

                // Attempt to fix errors
                Console.WriteLine("\nAttempting to fix grammar errors...");
                _doc.CheckGrammar();
                _doc.Save();

                // Second check - List remaining errors
                Console.WriteLine("\nRemaining Grammar Errors:");
                ListGrammarErrors();

                Console.WriteLine("\n=== Grammar Check Completed ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during grammar check: {ex.Message}");
            }
            finally
            {
                Cleanup();
            }
        }

        private void ListGrammarErrors()
        {
            if (_doc == null) return;

            var errorCount = _doc.GrammaticalErrors.Count;
            if (errorCount == 0)
            {
                Console.WriteLine("No grammatical errors found.");
                return;
            }

            Console.WriteLine($"Found {errorCount} grammatical error(s):");
            foreach (Microsoft.Office.Interop.Word.Range error in _doc.GrammaticalErrors)
            {
                Console.WriteLine($"- {error.Text}");
            }
        }

        private void Cleanup()
        {
            try
            {
                _doc?.Close();
                _wordApp?.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during cleanup: {ex.Message}");
            }
        }
    }
}