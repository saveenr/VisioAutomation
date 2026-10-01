using System.Collections.Generic;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.PowerShell
{
    // Helpers for cmdlet-binding tests that run whole scripts inside the session's runspace, so the
    // cmdlets go through PowerShell's parameter binder (see CmdletBindingTests for why that matters).
    // Everything a test needs is created inside the script and plain values are returned, which
    // keeps COM objects from crossing between the script and the test.
    public static class CmdletScriptExtensions
    {
        // Runs "body" in a fresh Visio document and returns what it writes to the pipeline.
        // The document is closed afterward without a save prompt.
        public static List<T> RunInNewDocument<T>(this VisioPSSession session, string body)
        {
            string script =
                "$doc = New-VisioDocument; " +
                "try { " + body + " } " +
                "finally { try { $doc.Saved = $true; $doc.Close() } catch { } }";
            return session.InvokeScriptStrict<T>(script);
        }

        // Runs "body" in a fresh Visio document where the script is expected to fail. Returns the
        // exception, which is flattened into its message chain by MessageOf. Fails the test if it succeeds.
        public static System.Exception ExpectFailureInNewDocument(this VisioPSSession session, string body)
        {
            try
            {
                session.RunInNewDocument<object>(body);
            }
            catch (System.Exception ex)
            {
                return ex;
            }

            MUT.Assert.Fail("Expected the script to fail: " + body);
            return null;
        }

        // The messages of an exception and everything inside it, so a test can look for a phrase
        // without caring how many layers PowerShell wrapped the error in.
        public static string MessageOf(System.Exception ex)
        {
            var parts = new List<string>();
            while (ex != null)
            {
                parts.Add(ex.GetType().Name + ": " + ex.Message);
                ex = ex.InnerException;
            }
            return string.Join(" | ", parts);
        }
    }
}
