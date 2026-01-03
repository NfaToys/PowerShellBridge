
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace PowerShellBridge
{
   // COM-visible interface for VBA to call
   [ComVisible(true)]
   [Guid("D0A94EF5-010B-4A9B-8B39-E14F7CC24003")]
   [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
   public interface IPowerShellBridge
   {
      void Initialize();
      void RunScriptFile(string scriptPath);
      void InvokeCommand(string commandText);
   }

   // COM-visible events interface for VBA to sink
   [ComVisible(true)]
   [Guid("C8B2091D-4F60-4D29-9A7A-2C16ED80C84B")]
   [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
   public interface IPowerShellBridgeEvents
   {
      void OutputReceived(string text);
      void ErrorReceived(string text);
      void WarningReceived(string text);
      void InformationReceived(string text);
      void DebugReceived(string text);
      void VerboseReceived(string text);
      void InvocationCompleted();
   }

   [ComVisible(true)]
   [Guid("69A860E7-DA53-4D1E-B413-08732BA2C1EE")]
   [ClassInterface(ClassInterfaceType.None)]
   [ComSourceInterfaces(typeof(IPowerShellBridgeEvents))]
   public class PowerShellRunner : IPowerShellBridge, IDisposable
   {
      private Runspace _runspace;
      private bool _initialized;

      public delegate void OutputReceivedEventHandler(string text);
      public event OutputReceivedEventHandler OutputReceived;

      public delegate void ErrorReceivedEventHandler(string text);
      public event ErrorReceivedEventHandler ErrorReceived;

      public delegate void WarningReceivedEventHandler(string text);
      public event WarningReceivedEventHandler WarningReceived;

      public delegate void InformationReceivedEventHandler(string text);
      public event InformationReceivedEventHandler InformationReceived;

      public delegate void DebugReceivedEventHandler(string text);
      public event DebugReceivedEventHandler DebugReceived;

      public delegate void VerboseReceivedEventHandler(string text);
      public event VerboseReceivedEventHandler VerboseReceived;

      public delegate void InvocationCompletedEventHandler();
      public event InvocationCompletedEventHandler InvocationCompleted;

      public void Initialize()
      {
         if (_initialized)
            return;

         _runspace = RunspaceFactory.CreateRunspace();
         _runspace.Open();
         _initialized = true;
      }

      public void RunScriptFile(string scriptPath)
      {
         EnsureInitialized();

         if (string.IsNullOrWhiteSpace(scriptPath))
            throw new ArgumentException("scriptPath is null or empty.", nameof(scriptPath));

         if (!File.Exists(scriptPath))
            throw new FileNotFoundException("PowerShell script file not found.", scriptPath);

         string scriptText = File.ReadAllText(scriptPath);
         ExecuteScriptInternal(scriptText);
      }

      public void InvokeCommand(string commandText)
      {
         EnsureInitialized();

         if (string.IsNullOrWhiteSpace(commandText))
            return; // ignore empty lines

         ExecuteScriptInternal(commandText);
      }

      private void ExecuteScriptInternal(string scriptText)
      {
         using (var ps = PowerShell.Create())
         {
            ps.Runspace = _runspace;

            // This allows line-by-line commands OR whole scripts
            ps.AddScript(scriptText);

            // Subscribe to streams BEFORE Invoke
            ps.Streams.Error.DataAdded += (s, e) =>
            {
               try
               {
                  var err = ps.Streams.Error[e.Index];
                  ErrorReceived?.Invoke(err.ToString());
               }
               catch { }
            };

            ps.Streams.Warning.DataAdded += (s, e) =>
            {
               try
               {
                  var w = ps.Streams.Warning[e.Index];
                  WarningReceived?.Invoke(w.ToString());
               }
               catch { }
            };

            ps.Streams.Information.DataAdded += (s, e) =>
            {
               try
               {
                  var info = ps.Streams.Information[e.Index];
                  InformationReceived?.Invoke(info.ToString());
               }
               catch { }
            };

            ps.Streams.Debug.DataAdded += (s, e) =>
            {
               try
               {
                  var d = ps.Streams.Debug[e.Index];
                  DebugReceived?.Invoke(d.ToString());
               }
               catch { }
            };

            ps.Streams.Verbose.DataAdded += (s, e) =>
            {
               try
               {
                  var v = ps.Streams.Verbose[e.Index];
                  VerboseReceived?.Invoke(v.ToString());
               }
               catch { }
            };

            // Invoke the pipeline
            var results = ps.Invoke();

            // Pipeline output objects
            foreach (var r in results)
            {
               try
               {
                  OutputReceived?.Invoke(r.ToString());
               }
               catch { }
            }

            // If there were errors not already surfaced, you could surface them here as well

            try
            {
               InvocationCompleted?.Invoke();
            }
            catch { }
         }
      }

      private void EnsureInitialized()
      {
         if (!_initialized)
            Initialize();
      }

      public void Dispose()
      {
         if (_runspace != null)
         {
            if (_runspace.RunspaceStateInfo.State == RunspaceState.Opened)
               _runspace.Close();
            _runspace.Dispose();
            _runspace = null;
         }
         _initialized = false;
      }
   }
}