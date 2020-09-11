/*
tool that will go through all emails in a mailbox (gmail) and dump 
all email addresses regardless of whether those addresses are saved as contacts.
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Threading;
using System.IO;
using MailKit.Search;
using MimeKit;
using System.Text.RegularExpressions;

namespace EmailIDsExtractor
{
    public class EmailConfiguration
    {
        public string ImapServer { get; set; }
        public int ImapPort { get; set; }
        public string ImapUsername { get; set; }
        public string ImapPassword { get; set; }
        public EmailConfiguration(string configPath)
        {
            XElement conn = XElement.Load(configPath);

            ImapServer = conn.Element("Server")?.Value;
            if (string.IsNullOrEmpty(ImapServer))
                throw new Exception("Server must be specified.");

            if (int.TryParse(conn.Element("Port")?.Value, out int port))
                ImapPort = port;
            else
                throw new Exception("Port must be specified.");

            ImapUsername = conn.Element("UserName")?.Value;
            if (string.IsNullOrEmpty(ImapUsername))
                throw new Exception("UserName must be specified.");

            ImapPassword = conn.Element("Password")?.Value;
            if (string.IsNullOrEmpty(ImapPassword))
                throw new Exception("Password must be specified.");
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;

            EmailConfiguration emailConfiguration;
            try
            {
                emailConfiguration = new EmailConfiguration(Path.Combine(path, "connection.config"));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("connection.config file could not be loaded. " + ex.Message);
                return;
            }

            string outputsPath = Path.Combine(path, "Outputs");
            Directory.CreateDirectory(outputsPath);
            int readCount = GetEmailCount(outputsPath);
            if (readCount > 0)
            {
                Console.WriteLine($"{readCount} emails were already processed earlier. These will be skipped");
            }

            ImapClient client = null;
            try
            {
                while (true)
                {
                    int prevCount = readCount;
                    string contactsPath = MakeFilePath(outputsPath, readCount);
                    var list = ProcessBatch(ref client, () => CreateClient(emailConfiguration), ref readCount);
                    if (prevCount == readCount)
                    {
                        list = ReadAddresses(contactsPath);
                        string backup = contactsPath + ".backup";
                        File.Move(contactsPath, backup);
                        WriteAddresses(list, contactsPath);
                        File.Delete(backup);
                        Console.WriteLine("Entire inbox processed");
                        Console.WriteLine("Press any key to close.");
                        Console.ReadKey();
                        return;
                    }
                    else
                    {
                        WriteAddresses(list, contactsPath);
                        File.Move(contactsPath, MakeFilePath(outputsPath, readCount));
                        Console.WriteLine($"{readCount} emails processed");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                Console.WriteLine("You can try running the program again. It will continue from the point of failure.");
                Console.WriteLine("Press any key to close.");
                Console.ReadKey();
            }
            finally
            {
                if (client != null) client.Dispose();
            }
        }

        const string FilePrefix = "contacts-";
        static string MakeFilePath(string folder, int count)
        {
            return Path.Combine(folder, $"{FilePrefix}{count:0}.csv");
        }
        static int GetEmailCount(string folderPath)
        {
            foreach (var path in Directory.GetFiles(folderPath, $"{FilePrefix}*.csv"))
            {
                if (int.TryParse(Path.GetFileNameWithoutExtension(path).Substring(FilePrefix.Length), out int count))
                    return count;
            }
            return 0;
        }

        static ImapClient CreateClient(EmailConfiguration configuration)
        {
            var client = new ImapClient();
            client.Connect(configuration.ImapServer, configuration.ImapPort, SecureSocketOptions.SslOnConnect);
            client.Authenticate(configuration.ImapUsername, configuration.ImapPassword);
            client.Inbox.Open(FolderAccess.ReadOnly);
            return client;
        }

        static HashSet<MailboxAddress> ProcessBatch(ref ImapClient client, Func<ImapClient> create, ref int readCount)
        {
            int retries = 3;
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    if (client == null)
                        client = create();

                    var list = new HashSet<MailboxAddress>(new AddressComparer());
                    var messages = client.Inbox.Fetch(readCount, readCount + 199, MessageSummaryItems.Headers);
                    int lastIndex = -1;
                    foreach (var msg in messages)
                    {
                        lastIndex = Math.Max(lastIndex, msg.Index);
                        ParseInto(msg.Headers["From"] ?? "", list);
                        ParseInto(msg.Headers["To"] ?? "", list);
                        ParseInto(msg.Headers["CC"] ?? "", list);
                    }
                    if (lastIndex >= 0)
                        readCount = lastIndex + 1;
                    return list;
                }
                catch when (i < retries - 1)
                {
                    if (client != null)
                        client.Dispose();
                }
            }
            return null;//This is not actually reachable
        }
        private static void ParseInto(string value, HashSet<MailboxAddress> list)
        {
            if (InternetAddressList.TryParse(value, out var addresses))
            {
                foreach (var address in addresses)
                {
                    if (address is MailboxAddress a)
                    {
                        list.Add(a);
                    }
                }
            }
        }

        private static void WriteAddresses(IEnumerable<MailboxAddress> addresses, string path)
        {
            if (!File.Exists(path))
                File.WriteAllText(path, $"Name,Email{Environment.NewLine}");

            File.AppendAllLines(path, addresses.Select(a => $"{Escape(a.Name)},{Escape(a.Address)}"));
        }

        private static string Escape(string s)
        {
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private static HashSet<MailboxAddress> ReadAddresses(string path)
        {
            var list = new HashSet<MailboxAddress>(new AddressComparer());
            var q = "\"";
            var regex = new Regex($@"{q}((?:[^{q}]|{q}{q})*){q},{q}((?:[^{q}]|{q}{q})*){q}");
            foreach (var line in File.ReadLines(path).Skip(1))
            {
                var match = regex.Match(line);
                var name = match.Groups[1].Value.Replace("\"\"", "\"");
                var address = match.Groups[2].Value.Replace("\"\"", "\"");
                list.Add(new MailboxAddress(name, address));
            }
            return list;
        }

        class AddressComparer : IEqualityComparer<MailboxAddress>
        {
            private static readonly StringComparer sc = StringComparer.OrdinalIgnoreCase;
            public bool Equals(MailboxAddress x, MailboxAddress y)
            {
                return sc.Equals(x.Name, y.Name) && sc.Equals(x.Address, y.Address);
            }

            public int GetHashCode(MailboxAddress obj)
            {
                return HashCode.Combine(sc.GetHashCode(obj.Name), sc.GetHashCode(obj.Address));
            }
        }
    }
}
