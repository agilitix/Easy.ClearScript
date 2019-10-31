using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;

namespace Easy.ClearScript
{
    internal class HostWindow : IHostWindow
    {
        public IntPtr OwnerHandle => IntPtr.Zero;

        public void EnableModeless(bool enable)
        {
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            using (var engine = new VBScriptEngine())
            {
                engine.HostWindow = new HostWindow(); // Parent window for the vbs ui calls.

                // Execute a string script.
                engine.Execute(@"
Set fso = CreateObject (""Scripting.FileSystemObject"")
Set stdout = fso.GetStandardStream (1)
Set stderr = fso.GetStandardStream (2)
stdout.WriteLine ""This will go to standard output.""
stderr.WriteLine ""This will go to error output.""

result = MsgBox(""Ready?"", vbQuestion + vbYesNo, ""Let's Go!"")
if (result = vbYes) then
    stdout.WriteLine(""yes="" & result)
else
    stdout.WriteLine(""no="" & result)
end if
");
            }

            using (var engine = new VBScriptEngine(WindowsScriptEngineFlags.EnableDebugging))
            {
                engine.HostWindow = new HostWindow();

                // Execute a file script.
                engine.DocumentSettings.AccessFlags = DocumentAccessFlags.EnableFileLoading;
                engine.ExecuteDocument("Sample.vbs");
            }

            Console.WriteLine();
            Console.Write("Exit:");
            Console.ReadKey();
        }
    }
}
