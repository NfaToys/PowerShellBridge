using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace PowerShellBridge
{
   // ============================
   // COM-visible contracts
   // ============================

   [ComVisible(true)]
   [Guid("D0A94EF5-010B-4A9B-8B39-E14F7CC24003")] // TODO: replace with your existing interface GUID
   [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
   public interface IPowerShellBridge
   {
      void Initialize();
      void InvokeCommand(string command);
      void RunScriptFile(string path);
      void SubmitInput(string text);
      void Dispose();
      object GetResult();
   }

   [ComVisible(true)]
   [Guid("C8B2091D-4F60-4D29-9A7A-2C16ED80C84B")] // TODO: replace with your existing events GUID
   [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
   public interface IPowerShellBridgeEvents
   {
      void OutputReceived(string text);
      void ErrorReceived(string text);
      void InvocationCompleted();
      void PromptRequested(string promptText);
      void ResultReady();
   }

   // ============================
   // COM-visible host class
   // ============================

   [ComVisible(true)]
   [Guid("69A860E7-DA53-4D1E-B413-08732BA2C1EE")] // TODO: replace with your existing class GUID
   [ClassInterface(ClassInterfaceType.None)]
   [ComSourceInterfaces(typeof(IPowerShellBridgeEvents))]
   public class PowerShellRunner : IPowerShellBridge, IDisposable
   {
      // --- Events exposed to VBA ---
      public delegate void OutputReceivedEventHandler(string text);
      public event OutputReceivedEventHandler OutputReceived;

      public delegate void ErrorReceivedEventHandler(string text);
      public event ErrorReceivedEventHandler ErrorReceived;

      public delegate void InvocationCompletedEventHandler();
      public event InvocationCompletedEventHandler InvocationCompleted;

      public delegate void PromptRequestedEventHandler(string promptText);
      public event PromptRequestedEventHandler PromptRequested;

      public delegate void ResultReadyEventHandler();
      public event ResultReadyEventHandler ResultReady;

      // --- Internal PowerShell state ---
      private Runspace _runspace;
      private readonly object _syncRoot = new object();

      // --- Result storage for pull-based output ---
      private object _lastResult;
      private readonly object _resultLock = new object();

      // --- Input coordination for Read-Host ---
      private readonly AutoResetEvent _inputReady = new AutoResetEvent(false);
      private readonly object _inputLock = new object();
      private string _pendingInput;

      // --- Host/HostUI ---
      private readonly BridgeHost _host;
      private readonly BridgeHostUI _hostUI;

      public PowerShellRunner()
      {
         _host = new BridgeHost(this);
         _hostUI = new BridgeHostUI(this);
      }

      // ============================
      // Public COM methods
      // ============================

      public void Initialize()
      {
         //System.Diagnostics.Debug.WriteLine("Initialize ENTER");
         lock (_syncRoot)
         {
            //if (_runspace != null &&
            //    _runspace.RunspaceStateInfo.State == RunspaceState.Opened)
            //   return;

            //var state = InitialSessionState.CreateDefault();
            //_runspace = RunspaceFactory.CreateRunspace(_host, state);
            //_runspace.Open();

            try
            {
               var state = InitialSessionState.CreateDefault();
               _runspace = RunspaceFactory.CreateRunspace(_host, state);
               _runspace.Open();
               //System.Diagnostics.Debug.WriteLine("Initialize SUCCESS");
            }
            catch (Exception ex)
            {
               //System.Diagnostics.Debug.WriteLine("Initialize EXCEPTION: " + ex);
               throw;
            }
         }
      }

      public void InvokeCommand(string command)
      {
         if (string.IsNullOrWhiteSpace(command))
            return;

         EnsureRunspace();

         ThreadPool.QueueUserWorkItem(_ =>
         {
            try
            {
               //System.Diagnostics.Debug.WriteLine("LAMBDA ENTER");
               //RaiseOutput("DEBUG: Lambda started");

               using (var ps = PowerShell.Create())
               {
                  //System.Diagnostics.Debug.WriteLine("PowerShell Object created");
                  //RaiseOutput("DEBUG: PowerShell.Create() OK");

                  ps.Runspace = _runspace;
                  //RaiseOutput("DEBUG: Runspace assigned");

                  ps.AddScript(command, useLocalScope: true);
                  //RaiseOutput("DEBUG: Script added");

                  AttachPipelines(ps);
                  //RaiseOutput("DEBUG: Pipelines attached");

                  var output = new PSDataCollection<PSObject>();
                  output.DataAdded += (sender, e) =>
                  {
                     var stream = (PSDataCollection<PSObject>)sender;
                     var obj = stream[e.Index];
                     //RaiseOutput(obj);
                  };

                  //RaiseOutput("DEBUG: Invoking...");
                  ps.Invoke(null, output);
                  //RaiseOutput("DEBUG: Invoke complete");

                  //foreach (var obj in output)
                  //   RaiseOutput(obj);

                  object result = null;

                  if (output.Count == 1)
                  {
                     result = output[0].BaseObject;
                  }
                  else if (output.Count > 1)
                  {
                     var list = new System.Collections.ArrayList();
                     foreach (var obj in output)
                        list.Add(obj.BaseObject);
                     result = list;
                  }

                  lock (_resultLock)
                  {
                     _lastResult = result;
                  }

                  // Notify VBA that the result is ready
                  try
                  {
                     ResultReady?.Invoke();
                  }
                  catch
                  {
                     // swallow COM event exceptions
                  }

                  RaiseInvocationCompleted();
               }
            }
            catch (Exception ex)
            {
               RaiseError("THREAD EXCEPTION: " + ex.ToString());
               RaiseInvocationCompleted();
            }
         }); ;
      }

      public void RunScriptFile(string path)
      {
         if (string.IsNullOrWhiteSpace(path))
            return;

         // Quote the path to handle spaces and quotes
         string command = $". '{path.Replace("'", "''")}'";
         InvokeCommand(command);
      }

      public void SubmitInput(string text)
      {
         lock (_inputLock)
         {
            _pendingInput = text ?? string.Empty;
         }
         _inputReady.Set();
      }

      public object GetResult()
      {
         lock (_resultLock)
         {
            return _lastResult;
         }
      }

      public void Dispose()
      {
         lock (_syncRoot)
         {
            if (_runspace != null)
            {
               try
               {
                  _runspace.Close();
                  _runspace.Dispose();
               }
               catch
               {
                  // ignore
               }
               _runspace = null;
            }
         }
      }

      // ============================
      // Internal plumbing
      // ============================

      private void EnsureRunspace()
      {
         Initialize();
      }

      private void AttachPipelines(PowerShell ps)
      {
         // Error stream
         ps.Streams.Error.DataAdded += (sender, e) =>
         {
            var stream = (PSDataCollection<ErrorRecord>)sender;
            var err = stream[e.Index];
            RaiseError(err.ToString());
         };
      }

      internal string WaitForInput(string promptText)
      {
         // Fire event to VBA
         RaisePromptRequested(promptText);

         // Wait until VBA calls SubmitInput
         _inputReady.WaitOne();

         lock (_inputLock)
         {
            return _pendingInput ?? string.Empty;
         }
      }

      // ============================
      // Event helpers
      // ============================

      internal void RaiseOutput(object obj)
      {
         string text = obj?.ToString() ?? string.Empty;

         try
         {
            //System.Diagnostics.Debug.WriteLine("OutputReceived?.Invoke(text) called text = " + text);
            OutputReceived?.Invoke(text);
            //System.Diagnostics.Debug.WriteLine("OutputReceived?.Invoke(text) done");
         }
         catch
         {
            // swallow COM event exceptions
         }
      }

      internal void RaiseError(string text)
      {
         try
         {
            ErrorReceived?.Invoke(text);
         }
         catch
         {
            // swallow COM event exceptions
         }
      }

      internal void RaiseInvocationCompleted()
      {
         try
         {
            InvocationCompleted?.Invoke();
         }
         catch
         {
            // swallow COM event exceptions
         }
      }

      internal void RaisePromptRequested(string promptText)
      {
         try
         {
            PromptRequested?.Invoke(promptText);
         }
         catch
         {
            // swallow COM event exceptions
         }
      }

      // ============================
      // Host + HostUI implementations
      // ============================

      private class BridgeHost : PSHost
      {
         private readonly PowerShellRunner _owner;
         private readonly Guid _instanceId = Guid.NewGuid();

         public BridgeHost(PowerShellRunner owner)
         {
            _owner = owner;
         }

         public override Guid InstanceId => _instanceId;

         public override string Name => "PowerShellBridgeHost";

         public override Version Version => new Version(1, 0);

         public override PSHostUserInterface UI => _owner._hostUI;

         public override CultureInfo CurrentCulture => CultureInfo.CurrentCulture;

         public override CultureInfo CurrentUICulture => CultureInfo.CurrentUICulture;

         public override void EnterNestedPrompt()
         {
         }

         public override void ExitNestedPrompt()
         {
         }

         public override void NotifyBeginApplication()
         {
         }

         public override void NotifyEndApplication()
         {
         }

         public override void SetShouldExit(int exitCode)
         {
         }
      }

      private class BridgeHostUI : PSHostUserInterface
      {
         private readonly PowerShellRunner _owner;
         private readonly BridgeRawUI _rawUI = new BridgeRawUI();

         public BridgeHostUI(PowerShellRunner owner)
         {
            _owner = owner;
         }

         public override PSHostRawUserInterface RawUI => _rawUI;

         public override void Write(string value)
         {
            _owner.RaiseOutput(value);
         }

         public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
         {
            _owner.RaiseOutput(value);
         }

         public override void WriteLine(string value)
         {
            _owner.RaiseOutput(value);
         }

         public override void WriteLine()
         {
            _owner.RaiseOutput(string.Empty);
         }

         public override void WriteErrorLine(string value)
         {
            _owner.RaiseError(value);
         }

         public override void WriteDebugLine(string message)
         {
            _owner.RaiseOutput("[DEBUG] " + message);
         }

         public override void WriteVerboseLine(string message)
         {
            _owner.RaiseOutput("[VERBOSE] " + message);
         }

         public override void WriteWarningLine(string message)
         {
            _owner.RaiseOutput("[WARN] " + message);
         }

         public override void WriteProgress(long sourceId, ProgressRecord record)
         {
            // Optional: raise a progress event if you want
         }

         public override string ReadLine()
         {
            const string defaultPrompt = "Input required:";
            return _owner.WaitForInput(defaultPrompt);
         }

         public override SecureString ReadLineAsSecureString()
         {
            string input = ReadLine();
            var ss = new SecureString();
            foreach (char c in input)
               ss.AppendChar(c);
            return ss;
         }

         public override Dictionary<string, PSObject> Prompt(
             string caption,
             string message,
             Collection<FieldDescription> descriptions)
         {
            var dict = new Dictionary<string, PSObject>();
            foreach (var fd in descriptions)
            {
               string promptText = $"{caption} {message} {fd.Name}".Trim();
               string val = _owner.WaitForInput(promptText);
               dict[fd.Name] = PSObject.AsPSObject(val);
            }
            return dict;
         }

         public override int PromptForChoice(
             string caption,
             string message,
             Collection<ChoiceDescription> choices,
             int defaultChoice)
         {
            var promptText = $"{caption} {message}";
            var val = _owner.WaitForInput(promptText + " (enter choice index)");

            if (int.TryParse(val, out int idx) &&
                idx >= 0 && idx < choices.Count)
               return idx;

            return defaultChoice;
         }

         public override PSCredential PromptForCredential(
             string caption,
             string message,
             string userName,
             string targetName)
         {
            string user = _owner.WaitForInput($"{caption} {message} User:");
            string pwd = _owner.WaitForInput($"{caption} {message} Password:");

            var ss = new SecureString();
            foreach (char c in pwd) ss.AppendChar(c);

            return new PSCredential(user, ss);
         }

         public override PSCredential PromptForCredential(
             string caption,
             string message,
             string userName,
             string targetName,
             PSCredentialTypes allowedCredentialTypes,
             PSCredentialUIOptions options)
         {
            return PromptForCredential(caption, message, userName, targetName);
         }
      }

      private class BridgeRawUI : PSHostRawUserInterface
      {
         private ConsoleColor _foreground = ConsoleColor.Gray;
         private ConsoleColor _background = ConsoleColor.Black;
         private Size _bufferSize = new Size(120, 50);
         private Coordinates _cursorPosition = new Coordinates(0, 0);
         private int _cursorSize = 1;
         private Size _windowSize = new Size(120, 50);
         private readonly Size _maxWindowSize = new Size(120, 50);
         private readonly Size _maxPhysicalWindowSize = new Size(120, 50);

         public override ConsoleColor BackgroundColor
         {
            get => _background; set => _background = value;
         }
         public override ConsoleColor ForegroundColor
         {
            get => _foreground; set => _foreground = value;
         }
         public override Size BufferSize
         {
            get => _bufferSize; set => _bufferSize = value;
         }
         public override Coordinates CursorPosition
         {
            get => _cursorPosition; set => _cursorPosition = value;
         }
         public override int CursorSize
         {
            get => _cursorSize; set => _cursorSize = value;
         }
         public override bool KeyAvailable => false;
         public override Size MaxPhysicalWindowSize => _maxPhysicalWindowSize;
         public override Size MaxWindowSize => _maxWindowSize;
         public override Coordinates WindowPosition
         {
            get => new Coordinates(0, 0); set
            {
            }
         }
         public override Size WindowSize
         {
            get => _windowSize; set => _windowSize = value;
         }
         public override string WindowTitle { get; set; } = "PowerShellBridge";

         public override void FlushInputBuffer()
         {
         }

         public override KeyInfo ReadKey(ReadKeyOptions options)
         {
            // Minimal implementation; no real key handling
            return new KeyInfo();
         }

         public override void SetBufferContents(Coordinates origin, BufferCell[,] contents)
         {
         }

         public override void SetBufferContents(Rectangle rectangle, BufferCell fill)
         {
         }

         public override BufferCell[,] GetBufferContents(Rectangle rectangle)
         {
            return new BufferCell[0, 0];
         }

         public override void ScrollBufferContents(Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill)
         {
         }
      }
   }
}