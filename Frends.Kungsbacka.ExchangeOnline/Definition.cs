#pragma warning disable 1591

using Newtonsoft.Json.Linq;
using System.ComponentModel.DataAnnotations;

namespace Frends.Kungsbacka.ExchangeOnline
{
    /// <summary>
    /// Required parameters for Connect task
    /// </summary>
    public class ConnectParameters
    {
        /// <summary>
        /// Exchange Online organisation (company.onmicrosoft.com)
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string Organization { get; set; }

        /// <summary>
        /// AppId for Azure AD service principal
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string AppId { get; set; }

        /// <summary>
        /// Certificate associated with service principal
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string Base64EncodedCertificate { get; set; }

        /// <summary>
        /// Password used to decrypt private key
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public string CertificatePassword { get; set; }
    }

    public class ConnectResult
    {
        /// <summary>
        /// ExchangeSession object
        /// </summary>
        public ExchangeSession Session { get; set; }
    }

    /// <summary>
    /// Disconnect does not give a result, but frends
    /// does not accept a task that returns void.
    /// </summary>
    public class DisconnectResult { }

    /// <summary>
    /// Required parameters for Disconnect task
    /// </summary>
    public class DisconnectParameters
    {
        /// <summary>
        /// Exchange session to disconnect
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public ExchangeSession Session { get; set; }
    }

    /// <summary>
    /// Required parameters for InvokeCommand task
    /// </summary>
    public class InvokeCommandParameters
    {
        /// <summary>
        /// Exchange session where the commnand is executed
        /// </summary>
        [DisplayFormat(DataFormatString = "Expression")]
        public ExchangeSession Session { get; set; }

        /// <summary>
        /// PowerShell command (i.g. Get-Mailbox)
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string Command { get; set; }
        
        /// <summary>
        /// Command parameters
        /// </summary>
        public CommandParameter[] Parameters { get; set; }
    }

    /// <summary>
    /// PowerShell command parameter
    /// </summary>
    public class CommandParameter
    {
        /// <summary>
        /// Parameter name
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string Name { get; set; }

        /// <summary>
        /// Parameter value
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public object Value { get; set; }
    }

    /// <summary>
    /// InvokeCommand result
    /// </summary>
    public class InvokeCommandResult
    {
        /// <summary>
        /// Result
        /// </summary>
        public JToken Result { get; set; }
    }
}
