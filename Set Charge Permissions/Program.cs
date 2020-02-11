using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using NDesk.Options;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.Globalization;

namespace Set_Charge_Permissions
{
  class Program
  {
    enum AuthScheme { ntlm, adfs };
    static void Main(string[] args)
    {
      try
      {
        var runStartTime = DateTime.Now;
        Console.WriteLine($"Run started {runStartTime}.");
        string groupPrefix = "";
        string chargesLstNm = "Charges";
        var authScheme = AuthScheme.ntlm;
        string adfsServer = null;
        string relyingParty = null;
        string userName = null;
        string domain = null;
        string pwd = null;
        var missingSecurityGroups = new HashSet<string>();
        DateTime? billingCycleStart = null;

        var options = new OptionSet(){
                    {"p|prefix_of_group=", v => groupPrefix = v}
                    ,{"c|cycle_start_date=", v => billingCycleStart = DateTime.ParseExact(v,"yyyy-MM-dd",CultureInfo.InvariantCulture)}
                    ,{"h|charges_list_name=", v => chargesLstNm = v}
                    ,{"s|auth_scheme=", v=> authScheme = (AuthScheme) Enum.Parse(typeof(AuthScheme), v)}
                    ,{"S|adfs_server=", v=> adfsServer = v}
                    ,{"P|relying_party=", v=> relyingParty = v}
                    ,{"u|username=", v=> userName = v}
                    ,{"D|domain=", v=> domain = v}
                    ,{"w|password=", v=> pwd = v}
                };
        List<String> extraArgs = options.Parse(args);
        if (billingCycleStart == null)
        {
          Console.Error.WriteLine($"Argument -c|--cycle_start_date is mandatory.");
          System.Environment.Exit(1);
        }
        ServicePointManager.ServerCertificateValidationCallback = MyCertHandler;
        ClientContext cc = null;
        switch (authScheme)
        {
          case AuthScheme.ntlm:
            {
              if (userName != null && pwd != null && domain != null)
              {
                OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();
                cc = am.GetNetworkCredentialAuthenticatedContext(extraArgs[0], userName, pwd, domain);
              }
              else
              {
                cc = new ClientContext(extraArgs[0]);
                cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
              }
              break;
            }
          case AuthScheme.adfs:
            {
              OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();
              cc = am.GetADFSUserNameMixedAuthenticatedContext(extraArgs[0], userName, pwd, domain, adfsServer, relyingParty);
              break;
            }
        }

        var query = new CamlQuery();
        // populate or update charge from consumption
        string viewXml = null;
        viewXml = $@"
<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='Cycle' />
         <Value Type='DateTime'>{((DateTime)billingCycleStart).ToString("yyyy-MM-dd")}</Value>
      </Eq>
   </Where>
</Query></View>";
        query.ViewXml = viewXml;
        cc.Load(cc.Web.RoleDefinitions);
        var gc = cc.Web.SiteGroups;
        cc.Load(gc);
        var readRD = cc.Web.RoleDefinitions.GetByName("Read");
        cc.ExecuteQuery();
        RoleDefinition restReadRD;
        try
        {
          restReadRD = cc.Web.RoleDefinitions.GetByName("Restricted Read");
          cc.ExecuteQuery();
        }
        catch
        {
          restReadRD = null;
        }
        var chargesLst = cc.Web.Lists.GetByTitle(chargesLstNm);
        var chargesLIC = chargesLst.GetItems(query);
        cc.Load(chargesLIC, items => items.Include(
            item => item["Account"]
            , item => item["HasUniqueRoleAssignments"]
            ));
        cc.ExecuteQuery();
        foreach (var chargeLI in chargesLIC)
        {
          if (chargeLI.HasUniqueRoleAssignments)
          {
            chargeLI.ResetRoleInheritance();
            cc.ExecuteQuery();
          }
          chargeLI.BreakRoleInheritance(true, false);
          bool match = false;
          foreach (var g in gc)
          {
            if (g.LoginName == (groupPrefix + chargeLI["Account"].ToString()))
            {
              var rdb = new RoleDefinitionBindingCollection(cc);
              rdb.Add(restReadRD != null ? restReadRD : readRD);
              chargeLI.RoleAssignments.Add(g, rdb);
              match = true;
              break;
            }
          }
          if (!match)
          {
            missingSecurityGroups.Add(chargeLI["Account"].ToString());
          }
          cc.ExecuteQuery();
        }
        if (missingSecurityGroups.Count > 0)
        {
          Console.WriteLine($"Following accounts miss corresponding sharepoint groups");
          foreach (String acct in missingSecurityGroups)
          {
            Console.WriteLine($"{acct}");
          }
        }
        var runEndTime = DateTime.Now;
        Console.WriteLine($"Run ended {runEndTime}, lasting {(runEndTime - runStartTime).TotalMinutes.ToString("0.00")} minutes.");
      }
      catch (Exception ex)
      {
        Console.WriteLine($"{ex}");
        throw;
      }
    }

    static bool MyCertHandler(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors error)
    {
      // Ignore errors
      return true;
    }
  }
}
