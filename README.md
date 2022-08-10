# Frends.Kungsbacka.ExchangeOnline

frends Task to connect to Exchange Online and issue commands. The task creates a PowerShell runspace and loads the ExchangeOnlineManagement module. The Connect task connects to Exchange Online using a service principal and certificate. Once connected you can issue all commands that are available in ExchangeOnlineManagement (Exo*) and the PowerShell remote session.

# Dependencies

The agent server needs to have the ExchangeOnlineManagement module installed for all users (Install-Module ExchangeOnlineManagement -Scope AllUsers -Force)

# Usage

Before you can issue commands you need to connect using the Connect task. The result contains a session object that you then use as an input parameter to all subsequent InvokeCommand tasks.
