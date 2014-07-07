using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Configuration;
using System.Net.Mail;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace SharepointListClient
{
    class Program
    {

        const int WAY_OVERDUE = 0;
        const int OVERDUE = 1;
        const int TODAY = 2;
        const int TOMORROW = 3;
        const int THIS_WEEK = 4;
    
        static void Main(string[] args)
        {
            try
            {
                // Starting with ClientContext, the constructor requires a URL to the 
                // server running SharePoint. 
                using (ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SiteUrl"]))
                {
                    CredentialCache cc = new CredentialCache();
                    NetworkCredential nc = new NetworkCredential(ConfigurationManager.AppSettings["Username"], ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);

                    cc.Add(new Uri(ConfigurationManager.AppSettings["SiteUrl"]), "Negotiate", nc);
                    context.Credentials = cc; 

                    Console.WriteLine("Connecting to Sharepoint: {0} ...", ConfigurationManager.AppSettings["SiteUrl"]);                    
                    // The SharePoint web at the URL.
                    Web web = context.Web;

                    // We want to retrieve the web's properties.
                    context.Load(web);
                    // Execute the query to the server.
                    context.ExecuteQuery();
                    Console.WriteLine("Connected to Sharepoint...");
                    
                    // Get the list properties. We need the EditURL
                    Console.WriteLine("Getting list: {0}...", ConfigurationManager.AppSettings["ListTitle"]);
                    var list = web.Lists.GetByTitle(ConfigurationManager.AppSettings["ListTitle"]);
                    context.Load(list);
                    context.ExecuteQuery();
                    var editFormUrl = list.DefaultEditFormUrl;

                    // Get items from the list that we will loop through and produce the email
                    Console.WriteLine("Getting items from list: {0}...", ConfigurationManager.AppSettings["ListTitle"]);
                    // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                    // so that it grabs all list items, regardless of the folder they are in. 
                    CamlQuery query = CamlQuery.CreateAllItemsQuery(Convert.ToInt32(ConfigurationManager.AppSettings["MaxItems"]));
                    var items = list.GetItems(query);

                    // Retrieve all items in the ListItemCollection from List.GetItems(Query). 
                    context.Load(items);
                    context.ExecuteQuery();

                    // This is the Edit URL for each item in the list
                    var editFormUri = new Uri(new Uri(ConfigurationManager.AppSettings["SiteUrl"]), new Uri(editFormUrl, UriKind.Relative));

                    Console.WriteLine("Process items...");

                    // Now process each task and send email
                    SendEmailReminder(context, items, editFormUri.ToString());
                }
                //Console.ReadLine();
            }
            catch (Exception x)
            {
                SendEmail(ConfigurationManager.AppSettings["FromEmail"], 
                    ConfigurationManager.AppSettings["ErrorTo"],
                    ConfigurationManager.AppSettings["ErrorTo"],
                    string.Format(ConfigurationManager.AppSettings["ErrorEmailSubject"], 
                        DateTime.Today.ToLongDateString()),
                    x.ToString());
            }
        }

        private static void SendEmailReminder(ClientContext context, ListItemCollection items, string editUrl)
        {
            var emailsToSend = new Dictionary<string, Dictionary<int,List<ListItem>>>();

            foreach (var listItem in items)
            {
                // let's use the filter criteria. Eg completed tasks are ignored. 
                var filterValue = (listItem[ConfigurationManager.AppSettings["FilterFieldName"]] ?? string.Empty).ToString();
                if (filterValue != ConfigurationManager.AppSettings["FilterFieldValue"])
                {
                    var dueDateString = (listItem[ConfigurationManager.AppSettings["DueDateFieldName"]] ?? string.Empty).ToString();
                    var modified = (listItem[ConfigurationManager.AppSettings["ModifiedFieldName"]] ?? string.Empty).ToString();
                    var assignedTo = listItem[ConfigurationManager.AppSettings["AssignedToFieldName"]] as FieldLookupValue;
                    var email = "";

                    // Find the email address of the user from username
                    if (assignedTo != null)
                    {
                        ListItem it = context.Web.SiteUserInfoList.GetItemById(assignedTo.LookupId);

                        context.Load(it);
                        context.ExecuteQuery();
                        email = Convert.ToString(it["EMail"]);
                    }

                    // From the due date, see whether the task is today, tomorrow, overdue etc
                    DateTime dueDate;
                    if (DateTime.TryParse(dueDateString, out dueDate) && !string.IsNullOrEmpty(email))
                    {
                        dueDate = dueDate.ToLocalTime();
                        if (!emailsToSend.ContainsKey(email))
                        {
                            emailsToSend.Add(email, new Dictionary<int, List<ListItem>>());
                            emailsToSend[email].Add(WAY_OVERDUE, new List<ListItem>());
                            emailsToSend[email].Add(OVERDUE, new List<ListItem>());
                            emailsToSend[email].Add(TODAY, new List<ListItem>());
                            emailsToSend[email].Add(TOMORROW, new List<ListItem>());
                            emailsToSend[email].Add(THIS_WEEK, new List<ListItem>());
                        }

                        // Today
                        if (dueDate.Date == DateTime.Today.Date)
                        {
                            emailsToSend[email][TODAY].Add(listItem);
                        }
                        // Overdue
                        else if (dueDate < DateTime.Today.Date)
                        {
                            DateTime modifiedDate = DateTime.Now;
                            if (DateTime.TryParse(modified, out modifiedDate))
                            {
                                var wayOverDueDelta = Convert.ToInt32(ConfigurationManager.AppSettings["WayOverdueDelta"]);
                                if (modifiedDate < dueDate.AddDays(-wayOverDueDelta))
                                    emailsToSend[email][WAY_OVERDUE].Add(listItem);
                                else
                                    emailsToSend[email][OVERDUE].Add(listItem);
                            }                            
                            else 
                            {
                                emailsToSend[email][OVERDUE].Add(listItem);
                            }
                        }
                        // Tomorrow
                        else if (dueDate.Date == DateTime.Today.AddDays(1).Date)
                        {
                            emailsToSend[email][TOMORROW].Add(listItem);
                        }
                        // This week
                        else if (dueDate.Date <= DateTime.Today.AddDays(7 - (int)DateTime.Today.DayOfWeek).Date)
                        {
                            emailsToSend[email][THIS_WEEK].Add(listItem);
                        }
                    }
                }
            }

            // Send email to each person
            foreach (string emailKey in emailsToSend.Keys)
            {
                Console.WriteLine("Send reminder for: {0}...", emailKey);
                SendEmail(emailKey, emailsToSend[emailKey], editUrl);
            }
        }

        static void SendEmail(string email, Dictionary<int, List<ListItem>> items, string editUrl)
        {
            // See if there's any reminder to send to the user. If not, don't send blank email
            var noEmail = true;
            foreach (var value in items.Values)
                if (value.Count > 0)
                    noEmail = false;

            if (noEmail)
            {
                Console.WriteLine("No immediate reminder for: {0}", email);
                return;
            }

            // Use the email templates and inject the task reminders inside them
            var emailTemplate = System.IO.File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "email_template.html"));
            var taskTemplate = System.IO.File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "task_template.html"));

            var tokens = new Dictionary<string, string>();
            tokens.Add("@{TODAY}", DateTime.Today.ToLongDateString());

            emailTemplate = ProcessTemplateBlock(emailTemplate, "@{HAS_WAY_OVERDUE_TASK}", "@{/HAS_WAY_OVERDUE_TASK}", items[WAY_OVERDUE].Count > 0);
            tokens.Add("@{WAY_OVERDUE_TASKS}", ConvertTasksToHtml(taskTemplate, items[WAY_OVERDUE], editUrl));
            
            emailTemplate = ProcessTemplateBlock(emailTemplate, "@{HAS_OVERDUE_TASK}", "@{/HAS_OVERDUE_TASK}", items[OVERDUE].Count > 0);
            tokens.Add("@{OVERDUE_TASKS}", ConvertTasksToHtml(taskTemplate, items[OVERDUE], editUrl));
            
            emailTemplate = ProcessTemplateBlock(emailTemplate, "@{HAS_TODAY_TASKS}", "@{/HAS_TODAY_TASKS}", items[TODAY].Count > 0);
            tokens.Add("@{TODAY_TASKS}", ConvertTasksToHtml(taskTemplate, items[TODAY], editUrl));
            
            emailTemplate = ProcessTemplateBlock(emailTemplate, "@{HAS_TOMORROW_TASKS}", "@{/HAS_TOMORROW_TASKS}", items[TOMORROW].Count > 0);
            tokens.Add("@{TOMORROW_TASKS}", ConvertTasksToHtml(taskTemplate, items[TOMORROW], editUrl));

            emailTemplate = ProcessTemplateBlock(emailTemplate, "@{HAS_THISWEEK_TASKS}", "@{/HAS_THISWEEK_TASKS}", items[THIS_WEEK].Count > 0);
            tokens.Add("@{THISWEEK_TASKS}", ConvertTasksToHtml(taskTemplate, items[THIS_WEEK], editUrl));

            var subject = string.Format(ConfigurationManager.AppSettings["EmailSubject"], DateTime.Today.ToLongDateString());
            var body = ReplaceTokens(emailTemplate, tokens);

            var filename = email.Replace('@', '_').Replace('.', '_') + ".html";
            System.IO.File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filename), body);

            SendEmail(ConfigurationManager.AppSettings["FromEmail"],
                email,
                ConfigurationManager.AppSettings["CcEmail"],
                subject,
                body);
        }

        private static string ProcessTemplateBlock(string template, string start, string end, bool keep)
        {
            int startPos = template.IndexOf(start);
            int endPos = template.IndexOf(end, startPos);

            if (keep)
                return template.Substring(0, startPos) 
                    + template.Substring(startPos + start.Length, endPos - (startPos + start.Length)) 
                    + template.Substring(endPos + end.Length);
            else 
                return template.Substring(0, startPos) + template.Substring(endPos + end.Length);
        }

        private static void SendEmail(string from, string to, string cc, string subject, string body)
        {
            var mail = new MailMessage(new MailAddress(from), new MailAddress(to));
            mail.CC.Add(new MailAddress(cc));
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;
            using (SmtpClient client = new SmtpClient())
            {
                client.Host = ConfigurationManager.AppSettings["SMTPServer"];
                client.Port = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);
                client.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["SMTPSSL"]);

                string username = ConfigurationManager.AppSettings["SMTPUserName"];
                string password = ConfigurationManager.AppSettings["SMTPPassword"];
                if (!string.IsNullOrEmpty(username))
                {
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential(username, password);
                }

                //validate the certificate
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                    { 
                        return true; 
                    };

                client.Send(mail);
            }
        }

        private static string ReplaceTokens(string template, Dictionary<string, string> tokens)
        {
            foreach (var key in tokens.Keys)
            {
                template = template.Replace(key, tokens[key]);
            }
            return template;
        }

        private static string ConvertTasksToHtml(string taskTemplate, List<ListItem> list, string editUrl)
        {
            var buffer = new StringBuilder();
            foreach (var listItem in list)
            {
                var taskValues = new Dictionary<string, string>();
                foreach (var key in listItem.FieldValues.Keys)
                {
                    DateTime attemptedDate;
                    var item = listItem[key];
                    if (item != null)
                    {
                        if (item is FieldLookupValue)
                            taskValues.Add(key, ((FieldLookupValue)item).LookupValue);
                        else if (DateTime.TryParse(item.ToString(), out attemptedDate))
                            taskValues.Add(key, DateTime.Parse(item.ToString()).ToLongDateString());
                        else
                            taskValues.Add(key, item.ToString());
                    }
                }
                
                // Find each @{xx} and see if xx appears in the listitem
                buffer.Append(Regex.Replace(taskTemplate, "@{([^}]*)}", new MatchEvaluator(match => 
                    {
                        var key = match.Groups[1].Value;
                        if (key == "EditUrl")
                            return editUrl + "?ID=" + listItem.Id;
                        else if (key == "ShortBody")
                        {
                            var body = (listItem["Body"] ?? string.Empty).ToString();
                            body = Regex.Replace(body, "<[^>]*>", "");
                            return (body.Length > 300 ? body.Substring(0, 300) + "..." : body);
                        }
                        else
                            if (taskValues.ContainsKey(key))
                                return taskValues[key];
                            else
                            {
                                Console.WriteLine("Unknown key:", key);
                                Console.WriteLine("Available keys:");
                                foreach (var name in listItem.FieldValues.Keys)
                                    Console.WriteLine(name);
                                return "Unknown key:" + key;
                            }
                    })));                
            }

            if (buffer.Length == 0)
                buffer.Append("No tasks.");

            return buffer.ToString();
        }
    }
}
