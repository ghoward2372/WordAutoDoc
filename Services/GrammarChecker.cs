using Microsoft.Office.Interop.Word;
using System;

namespace DocumentProcessor.Services
{
    public class GrammarChecker
    {
        private readonly Application _wordApp;
        private readonly Document _doc;

        public GrammarChecker(string filePath)
        {
            try
            {
                _wordApp = new Application { Visible = false };
                _doc = _wordApp.Documents.Open(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing Word application: {ex.Message}");
                throw;
            }
        }

        public void CheckAndFixGrammar()
        {
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