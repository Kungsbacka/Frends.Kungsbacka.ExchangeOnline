using System;
using System.Collections;
using System.Security.Cryptography.X509Certificates;
using Newtonsoft.Json.Linq;

#pragma warning disable 1591

namespace Frends.Kungsbacka.ExchangeOnline
{
    public static class Exo
    {
        /// <summary>
        /// Connect to Exchange Online and create a session object that 
        /// can be used by InvokeCommand.
        /// </summary>
        /// <param name="input">Required parameters to connect</param>
        /// <returns>{ConnectionResult}</returns>
        public static ConnectResult Connect(ConnectParameters input)
        {
            byte[] certBytes = Convert.FromBase64String(input.Base64EncodedCertificate);
            X509Certificate2 certificate = new(certBytes, input.CertificatePassword, X509KeyStorageFlags.PersistKeySet);
            ExchangeSession session = new(input.Organization, input.AppId, certificate);
            session.Connect();
            return new ConnectResult() { Session = session };
        }

        /// <summary>
        /// Disconnect from Exchange Online. It's a good practice to disconnect
        /// a session when it's not needed anymore since the number of connections
        /// to Exchange Online are limited.
        /// </summary>
        /// <param name="input">Required parameters to disconnect</param>
        /// <returns>(Nothing)</returns>
        public static DisconnectResult Disconnect(DisconnectParameters input)
        {
            input.Session.Disconnect();
            input.Session.Dispose();
            input.Session = null;
            return new DisconnectResult();
        }

        /// <summary>
        /// Execute a command in the supplied session
        /// </summary>
        /// <param name="input"></param>
        /// <returns>{InvokeCommandResult}</returns>
        public static InvokeCommandResult InvokeCommand(InvokeCommandParameters input)
        {
            IDictionary parameters = new Hashtable(input.Parameters.Length);
            foreach (CommandParameter param in input.Parameters)
            {
                parameters[param.Name] = param.Value;
            }

            JToken result = input.Session.InvokeCommand(input.Command, parameters);
            return new InvokeCommandResult() { Result = result };
        }
    }
}
