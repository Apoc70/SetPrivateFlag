// SetPrivateFlag
//
// Authors: Torsten Schlopsnies, Thomas Stensitzki
//
// Published under MIT license
//
// Read more in the following blog post: https://www.granikos.eu/en/justcantgetenough/PostId/379/set-mailbox-item-private-flag
//
// Find more Exchange community scripts at: http://scripts.granikos.eu
// Please report issues or feature request here: https://github.com/Apoc70/SetPrivateFlag/issues 
//
// Version 1.0.0.0 | Published 2018-01-11

using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Net;

// Configure log4net using the .config file
[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace SetPrivateFlag
{
    internal class Program
    {
        private static FindFoldersResults findFolders;
        private static FindItemsResults<Item> findResults;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Main program section
        /// Handling all validation and code logic
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                // getting all arguments from the command line
                var arguments = new UtilityArguments(args);

                if (arguments.Help)
                {
                    DisplayHelp();
                    Environment.Exit(0);
                }

                log.Info("Application started");

                string Mailbox = arguments.Mailbox;
                string Subject = arguments.Subject;

                if ((Mailbox == null) || (Mailbox.Length == 0))
                {
                    string Message = "No mailbox is given. Use -help to refer to the usage.";

                    if (log.IsWarnEnabled)
                    {
                        log.Warn(Message);
                    }
                    else
                    {
                        Console.WriteLine(Message);
                    }

                    DisplayHelp();
                    Environment.Exit(1);
                }

                if ((Subject == null) || (Subject.Length == 0))
                {
                    string Message = "No subject filter is given. Use -help to refer to the usage.";

                    if (log.IsWarnEnabled)
                    {
                        log.Warn(Message);
                    }
                    else
                    {
                        Console.WriteLine(Message);
                    }

                    DisplayHelp();
                    Environment.Exit(1);
                }

                // Log all arguments if DEBUG is set in xml
                log.Debug("Parsing arguments: ");
                log.Debug(string.Format("mailbox: {0}", arguments.Mailbox));
                log.Debug(string.Format("subject: {0}", arguments.Subject));
                log.Debug(string.Format("Help: {0}", arguments.Help));
                log.Debug(string.Format("noconfirmation: {0}", arguments.noconfirmation));
                log.Debug(string.Format("logonly: {0}", arguments.LogOnly));
                log.Debug(string.Format("impersonate: {0}", arguments.impersonate));
                log.Debug(string.Format("allowredirection: {0}", arguments.AllowRedirection));

                if (arguments.User != null)
                {
                    log.Debug(string.Format("User: {0}", arguments.User));
                }

                if (arguments.Password != null)
                {
                    log.Debug("Password: is set");
                }

                log.Debug(string.Format("ignorecertificate: {0}", arguments.IgnoreCertificate));
                if (arguments.URL != null)
                {
                    log.Debug(string.Format("Server URL: {0}", arguments.URL));
                }
                else
                {
                    log.Debug("Server URL: using AutoDiscover");
                }

                // Check if we need to ignore certificate errors
                // need to be set before the service is created
                if (arguments.IgnoreCertificate)
                {
                    log.Warn("Ignoring SSL error because option -ignorecertificate is set");
                    ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                }

                // create the EWS service
                ExchangeService ExService;
                // connect to the server
                if (arguments.URL != null)
                {
                    ExService = ConnectToExchange(Mailbox, arguments.URL, arguments.User, arguments.Password, arguments.impersonate);
                }
                else
                {
                    ExService = ConnectToExchange(Mailbox, arguments.AllowRedirection, arguments.User, arguments.Password, arguments.impersonate);
                }

                if (log.IsInfoEnabled) log.Info("Service created.");

                // find all folders (under MsgFolderRoot)
                List<Folder> FolderList = Folders(ExService);

                // now try to find all items that are marked as "private"
                for (int i = FolderList.Count - 1; i >= 0; i--)
                {
                    if (log.IsDebugEnabled) log.Debug(string.Format("Processing folder \"{0}\"", GetFolderPath(ExService, FolderList[i].Id)));

                    if (log.IsDebugEnabled) log.Debug(string.Format("ID: {0}", FolderList[i].Id));

                    List<Item> Results = PrivateItems(FolderList[i], Subject);

                    if (Results.Count > 0)
                    {
                        if (log.IsInfoEnabled) log.Info(string.Format("Private items in folder: {0}", Results.Count));
                    }

                    foreach (var Result in Results)
                    {
                        if (Result is EmailMessage)
                        {
                            if (log.IsInfoEnabled)
                            {
                                if (log.IsInfoEnabled) log.Info(string.Format("Elements found. Folder: \"{0}\" ", GetFolderPath(ExService, FolderList[i].Id)));
                                if (log.IsInfoEnabled) log.Info(string.Format("Subject: \"{0}\"", Result.Subject));
                                if (log.IsDebugEnabled) log.Debug(string.Format("ID of the item: {0}", Result.Id));

                            }
                            else
                            {
                                Console.WriteLine("Element found. Folder: {0}", GetFolderPath(ExService, FolderList[i].Id));
                                Console.WriteLine("Subject: {0}", Result.Subject);
                            }
                            if (!(arguments.noconfirmation))
                            {
                                if (!(arguments.LogOnly))
                                {
                                    Console.WriteLine(string.Format("Change to private? (Y/N) (Folder: {0} - Subject {1})", GetFolderPath(ExService, FolderList[i].Id), Result.Subject));
                                    string Question = Console.ReadLine();

                                    if (Question == "y" || Question == "Y")
                                    {
                                        log.Info("Change the item? Answer: Yes.");
                                        ChangeItem(Result);
                                    }
                                }
                            }
                            else
                            {
                                if (!(arguments.LogOnly))
                                {
                                    string Message = "Changing item without confirmation because -noconfirmation is true.";
                                    if (log.IsInfoEnabled)
                                    {
                                        log.Info(Message);
                                    }
                                    else
                                    {
                                        Console.WriteLine(Message);
                                    }
                                    ChangeItem(Result);
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                DisplayHelp();
                Environment.Exit(1);
            }
        }

        /// <summary>
        /// Change a mailbox item by updating MAPI ExtendedProperty 0x36
        /// </summary>
        /// <param name="Message">The message to update</param>
        public static void ChangeItem(Item Message)
        {
            // do we have the extended properties?
            if (Message.ExtendedProperties.Count > 0)
            {
                try
                {
                    var extendedPropertyDefinition = new ExtendedPropertyDefinition(0x36, MapiPropertyType.Integer);
                    int extendedPropertyindex = 0;

                    foreach (var extendedProperty in Message.ExtendedProperties)
                    {
                        if (extendedProperty.PropertyDefinition == extendedPropertyDefinition)
                        {
                            if (log.IsInfoEnabled)
                            {
                                log.Info(string.Format("Try to alter the message: {0}", Message.Subject));
                            }
                            else
                            {
                                Console.WriteLine("Try to alter the message: {0}", Message.Subject);
                            }

                            // Set the value of the extended property to 0 (which is Sensitivity normal, 2 would be private)
                            Message.ExtendedProperties[extendedPropertyindex].Value = 2;

                            // Update the item on the server with the new client-side value of the target extended property
                            Message.Update(ConflictResolutionMode.AlwaysOverwrite);
                        }
                        extendedPropertyindex++;
                    }
                }
                catch (Exception ex)
                {
                    log.Error("Error on update the item. Error message:", ex);
                }
                if (log.IsInfoEnabled)
                {
                    log.Info("Successfully changed");
                }
                else
                {
                    Console.WriteLine("Successfully changed");
                }
            }
        }

        /// <summary>
        /// Connect to Exchange using AutoDiscover for the given email address
        /// </summary>
        /// <param name="MailboxID">The users email address</param>
        /// <returns>Exchange Web Service binding</returns>
        public static ExchangeService ConnectToExchange(string MailboxID, bool allowredirection, string User, string Password, bool Impersonisation)
        {
            log.Info(string.Format("Connect to mailbox {0}", MailboxID));
            try
            {
                var service = new ExchangeService();

                if ((User == null) | (Password == null))
                {
                    service.UseDefaultCredentials = true;
                }
                else
                {
                    service.Credentials = new WebCredentials(User, Password);
                }

                if (allowredirection)
                {
                    service.AutodiscoverUrl(MailboxID, RedirectionCallback);
                }
                else
                {
                    service.AutodiscoverUrl(MailboxID);
                }
                if (Impersonisation)
                {
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, MailboxID);
                }
                return service;
            }
            catch (Exception ex)
            {
                log.Error("Connection to mailbox failed", ex);
                Environment.Exit(3);
            }
            // We will not reach this point, so null is ok here
            return null;
        }

        /// <summary>
        /// Connect to Exchange using AutoDiscover for the given email address
        /// </summary>
        /// <param name="MailboxID">The users email address</param>
        /// <returns>Exchange Web Service binding</returns>
        public static ExchangeService ConnectToExchange(string MailboxID, string URL, string User, string Password, bool Impersonisation)
        {
            log.Info(string.Format("Connect to mailbox {0}", MailboxID));
            try
            {
                var service = new ExchangeService();

                if ((User == null) | (Password == null))
                {
                    service.UseDefaultCredentials = true;
                }
                else
                {
                    service.Credentials = new WebCredentials(User, Password);
                }
                service.Url = new Uri(URL);
                if (Impersonisation)
                {
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, MailboxID);
                }

                return service;
            }
            catch (Exception ex)
            {
                log.Error("Connection to mailbox failed", ex);
                Environment.Exit(3);
            }
            // We will not reach this point, so null is ok here
            return null;
        }

        /// <summary>
        /// Get a single mailbox folder path
        /// </summary>
        /// <param name="service">The active EWs connection</param>
        /// <param name="ID">The mailbox folder Id</param>
        /// <returns>A string containing the current mailbox folder path</returns>
        public static string GetFolderPath(ExchangeService service, FolderId ID)
        {
            try
            {
                var FolderPathProperty = new ExtendedPropertyDefinition(0x66B5, MapiPropertyType.String);

                PropertySet psset1 = new PropertySet(BasePropertySet.FirstClassProperties);
                psset1.Add(FolderPathProperty);

                Folder FolderwithPath = Folder.Bind(service, ID, psset1);
                Object FolderPathVal = null;

                if (FolderwithPath.TryGetProperty(FolderPathProperty, out FolderPathVal))
                {
                    // because the FolderPath contains characters we don't want, we need to fix it
                    string FolderPathTemp = FolderPathVal.ToString();
                    if (FolderPathTemp.Contains("￾"))
                    {
                        return FolderPathTemp.Replace("￾", "\\");
                    }
                    else
                    {
                        return FolderPathTemp;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Failed to get folder path", ex);
            }

            return "";
        }

        /// <summary>
        /// Find all folders under MsgRootFolder
        /// </summary>
        /// <param name="service"></param>
        /// <returns>Result of a folder search operation</returns>
        public static List<Folder> Folders(ExchangeService service)
        {
            // try to find all folder that are unter MsgRootFolder
            int pageSize = 100;
            int pageOffset = 0;
            bool moreItems = true;
            var view = new FolderView(pageSize, pageOffset);
            var resultFolders = new List<Folder>();

            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);

            // we define the seacht filter here. Find all folders which hold more than 0 elements
            SearchFilter searchFilter = new SearchFilter.IsGreaterThan(FolderSchema.TotalCount, 0);
            view.Traversal = FolderTraversal.Deep;

            while (moreItems)
            {
                try
                {
                    findFolders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, view);
                    moreItems = findFolders.MoreAvailable;

                    foreach (var folder in findFolders)
                    {
                        resultFolders.Add(folder);
                    }
                    // if more folders than the offset is aviable we need to page
                    if (moreItems) view.Offset += pageSize;
                }
                catch (Exception ex)
                {
                    log.Error("Failed to fetch folders.", ex);
                    moreItems = false;
                    Environment.Exit(3);
                }
            }
            return resultFolders;
        }

        /// <summary>
        /// Find items having a ExtendedPropertyDefinition 0x36 
        /// </summary>
        /// <param name="MailboxFolder">The mailbox folder to search</param>
        /// <returns>Items of an item search operation</returns>
        public static List<Item> PrivateItems(Folder MailboxFolder, string Subject)
        {
            int pageSize = 100;
            int pageOffset = 0;
            bool moreItems = true;
            var resultItems = new List<Item>();

            var extendedPropertyDefinition = new ExtendedPropertyDefinition(0x36, MapiPropertyType.Integer);
            SearchFilter searchFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, Subject);

            var view = new ItemView(pageSize, pageOffset);
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Sensitivity, ItemSchema.Subject, extendedPropertyDefinition);
            view.Traversal = ItemTraversal.Shallow;

            while (moreItems)
            {
                try
                {
                    findResults = MailboxFolder.FindItems(searchFilter, view);
                    moreItems = findResults.MoreAvailable;

                    foreach (var Found in findResults)
                    {
                        resultItems.Add(Found);
                    }

                    // if more folders than the offset is aviable we need to page
                    if (moreItems) view.Offset += pageSize;
                }
                catch (Exception ex)
                {
                    log.Error("Failed to fetch items.", ex);
                    moreItems = false;
                    Environment.Exit(3);
                }
            }
            return resultItems;
        }

        // Redirection Handler if -allowredirect is set
        public static bool RedirectionCallback(string url) =>
            // Return true if the URL is an HTTPS URL.
            url.ToLower().StartsWith("https://");

        /// <summary>
        /// Just some plain help message
        /// </summary>
        public static void DisplayHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("SetPrivateFlag.exe -mailbox \"user@example.com\" -subject \"[private]\" [-logonly] [-noconfirmation] [-ignorecertificate] [-url \"https://server/EWS/Exchange.asmx\"] [-user user@example.com] [-password Pa$$w0rd] [-impersonate]");
        }
    }
}
