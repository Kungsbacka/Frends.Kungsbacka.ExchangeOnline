using System;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security.Cryptography.X509Certificates;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Frends.Kungsbacka.ExchangeOnline
{
    /// <summary>
    /// Supplies an Exchange Online PowerShell remoting session
    /// </summary>
    public class ExchangeSession : IDisposable
    {
        private Runspace _runspace;
        private bool _disposed;
        private bool _connected;
        private readonly string _organization;
        private readonly string _appId;
        private readonly X509Certificate2 _certificate;

        /// <summary>
        /// Creates a new Exchange Online PowerShell session
        /// </summary>
        /// <param name="organization"></param>
        /// <param name="appId"></param>
        /// <param name="certificate"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public ExchangeSession(string organization, string appId, X509Certificate2 certificate)
        {
            _organization = organization ?? throw new ArgumentNullException(nameof(organization));
            _appId = appId ?? throw new ArgumentNullException(nameof(appId));
            _certificate = certificate ?? throw new ArgumentNullException(nameof(certificate));
            _connected = false;
        }

        private void CreateRunspaceAndConnect()
        {
            if (_connected)
            {
                return;
            }

            if (_runspace != null)
            {
                Dispose();
                throw new InvalidOperationException("ExchangeSession object is in an invalid state.");
            }

            string script = @"
                param ($Organization, $AppId, $Certificate)
                Set-ExecutionPolicy RemoteSigned
                Import-Module ExchangeOnlineManagement -MinimumVersion '2.0.5'
                Connect-ExchangeOnline -Organization $Organization -AppId $AppId -Certificate $Certificate -ShowBanner:$false -ShowProgress $false
            ";

            _runspace = RunspaceFactory.CreateRunspace();
            _runspace.Open();
            using PowerShell ps = PowerShell.Create();
            ps.Runspace = _runspace;
            ps.AddScript(script)
                .AddParameter("Organization", _organization)
                .AddParameter("AppId", _appId)
                .AddParameter("Certificate", _certificate);
            ps.Invoke();
            if (ps.HadErrors)
            {
                throw new AggregateException("Error creating Exchange Online PowerShell session",
                    ps.Streams.Error.Select(e => e.Exception).ToArray());
            }

            _connected = true;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException("PowerShell runspace is disposed. You must create a new ExchangeSession object.");
            }
        }

        /// <summary>
        /// Executes a command in a runspace connected to Exchange Online
        /// via PowerShell remoting. 
        /// </summary>
        /// <param name="command">Command to execute (single command)</param>
        /// <param name="parameters">Parameters</param>
        /// <returns>JToken</returns>
        public JToken InvokeCommand(string command, System.Collections.IDictionary parameters)
        {
            ThrowIfDisposed();
            if (!_connected)
            {
                throw new InvalidOperationException("You must call Connect() before calling any other method.");
            }

            using PowerShell ps = PowerShell.Create();
            ps.Runspace = _runspace;
            ps.AddCommand(command)
                    .AddParameters(parameters)
                .AddCommand("ConvertTo-Json")
                    .AddParameter("Depth", 10)
                    .AddParameter("Compress")
                    .AddParameter("ErrorAction", "SilentlyContinue");
            var result = ps.Invoke();
            StringBuilder sb = new();
            foreach (PSObject obj in result)
            {
                sb.Append(obj.ToString());
            }
            if (sb.Length > 0)
            {
                return JToken.Parse(sb.ToString());

            }
            return new JObject();
        }

        /// <summary>
        /// Connect to Exchange Online
        /// </summary>
        public void Connect()
        {
            ThrowIfDisposed();
            if (!_connected)
            {
                CreateRunspaceAndConnect();
            }
        }

        /// <summary>
        /// Disconnect from Exchange Online
        /// </summary>
        public void Disconnect()
        {
            ThrowIfDisposed();
            if (!_connected)
            {
                return;
            }

            using PowerShell ps = PowerShell.Create();
            ps.Runspace = _runspace;
            ps.AddCommand("Disconnect-ExchangeOnline")
                .AddParameter("Confirm", false);
            ps.Invoke();
            Dispose();
        }

        
        /// <summary>
        /// IDispose implementation
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_runspace != null)
                    {
                        _runspace.Dispose();
                    }
                }
                _disposed = true;
            }
        }

        /// <summary>
        /// IDispose implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
