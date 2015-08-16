using Microsoft.SharePoint.Client;
using System;

namespace SitePolicyEmailChanger
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteCollectionUrl = "";
            string relativeSiteUrl = "";

            // Get the site and web info
            Console.WriteLine("Site Policy E-mail Changer");
            Console.WriteLine("Enter the Site Collection URL: ");
            siteCollectionUrl = Console.ReadLine();
            Console.WriteLine("Enter the relative Site URL: ");
            relativeSiteUrl = Console.ReadLine();

            // Return the currently applied Site Policy
            ClientContext context = new ClientContext(siteCollectionUrl);
            Site site = context.Site;
            Web web = site.OpenWeb(relativeSiteUrl);
            ProjectPolicy policy = ProjectPolicy.GetCurrentlyAppliedProjectPolicyOnWeb(context, web);
            context.Load(policy,
                         p => p.Name,
                         p => p.Description,
                         p => p.EmailSubject,
                         p => p.EmailBody,
                         p => p.EmailBodyWithTeamMailbox);
            context.ExecuteQuery();

            // Display the current Site Policy properties and pause
            Console.WriteLine(String.Format("Policy Name is: {0}", policy.Name));
            Console.WriteLine(String.Format("Policy Description is: {0}", policy.Description));
            Console.WriteLine(String.Format("Policy E-mail Subject is: {0}", policy.EmailSubject));
            Console.WriteLine(String.Format("Policy E-mail Body is: {0}", policy.EmailBody));
            Console.WriteLine(String.Format("Policy E-mail Body (with Site Mailbox) is: {0}", policy.EmailBodyWithTeamMailbox));
            Console.WriteLine();
            Console.ReadLine();

            // Edit the Site Policy E-mail properties
            policy.EmailSubject = "Contoso Site Deletion Notice";
            policy.EmailBody = "The Contoso site <!--{SiteUrl}--> is set to expire on <!--{SiteDeleteDate}-->. If you have any questions or concerns, please contact your admin.";
            policy.EmailBodyWithTeamMailbox = "The Contoso site <!--{SiteUrl}--> associated with Site Mailbox <!--{TeamMailboxID}--> is set to expire on <!--{SiteDeleteDate}-->. If you have any questions or concerns, please contact your admin.";
            policy.SavePolicy();
            context.ExecuteQuery();

            // Refetch the edited Site Policy from the server
            policy = ProjectPolicy.GetCurrentlyAppliedProjectPolicyOnWeb(context, web);
            context.Load(policy,
                         p => p.Name,
                         p => p.Description,
                         p => p.EmailSubject,
                         p => p.EmailBody,
                         p => p.EmailBodyWithTeamMailbox);
            context.ExecuteQuery();

            // Display the new Site Policy properties and pause
            Console.WriteLine(String.Format("Policy Name is: {0}", policy.Name));
            Console.WriteLine(String.Format("Policy Description is: {0}", policy.Description));
            Console.WriteLine(String.Format("Policy E-mail Subject is NOW: {0}", policy.EmailSubject));
            Console.WriteLine(String.Format("Policy E-mail Body is NOW : {0}", policy.EmailBody));
            Console.WriteLine(String.Format("Policy E-mail Body (with Site Mailbox) is NOW: {0}", policy.EmailBodyWithTeamMailbox));
            Console.ReadLine();
        }
    }
}