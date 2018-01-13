## Synopsis

SetPrivateFlags.exe will SET the _private_ flag on all messages matching a subject substring.

See [https://github.com/Apoc70/RemovePrivateFlag](https://github.com/Apoc70/RemovePrivateFlag), if you want to _remove_ private message flags.

## Why should I use this

When you migrate from an alternative email server solution (e.g. Lotus Notes) and a mailbox contains confidential messages, you might want to mark such message as _private_ after content migration.

## Installation

Simply copy all files to a folder location where you are allowed to execute it and the Exchange servers are reachable.

## Requirements

* Exchange Server 2013/2016 (Tested with CU15, maybe it will work with Exchange 2007/2010 as well)
* _Application Impersonation Rights_ if you want to change items on other mailboxes than yours
* Microsoft.Exchange.WebServices.dll, log4net.dll (are provided in the repository and also in the binaries)

## Usage

```
SetPrivateFlags.exe -mailbox user@domain.com -subject "[private]"
```

Search the mailbox of user@domain.com for all message containing _[private]_ in the subject text and ask for changing _each_ item

```
SetPrivateFlags.exe -mailbox user@domain.com -subject "[private]" -noconfirmation
```

Search the mailbox of user@domain.com for all message containing _[private]_ in the subject text and all items are altered without asking for confirmation

# Parameters

* mandatory: -mailbox user@domain.com

Mailbox which you want to alter

* optional: -logonly

Items will only be logged

* optional: -noconfirmation

Messages will be set to normal without confirmation

* optional: -ignorecertificate

Ignore certificate errors. Interesting if you connect to a lab config with self signed certificate

* optional: -impersonate

If you want to alter a other mailbox than yours set this parameter

* optional: -user user@domain.com

If set together with -password this credentials would be used. Elsewhere the credentials from your session will be used

* optional: -password "Pa$$w0rd"
* optional: -url "https://server/EWS/Exchange.asmx"

If you set an specific URL this URL will be used instead of autodiscover. Should be used with -ignorecertificate if your CN is not in the certficate

* optional: -allowredirection

If your autodiscover redirects you the default behaviour is to quit the connection. With this parameter you will be connected anyhow

## License

MIT License
