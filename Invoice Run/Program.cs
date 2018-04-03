using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Diagnostics;
using NDesk.Options;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.Text.RegularExpressions;
using Itenso.TimePeriod;
using System.Globalization;

namespace Invoice_Run
{
  class Program
  {
    const string evtLogSrc = "BillEase";
    static void Main(string[] args)
    {
      if (!EventLog.SourceExists(evtLogSrc))
      {
        EventLog.CreateEventSource(evtLogSrc, "Application");
      }
      try
      {
        var runStartTime = DateTime.Now;
        EventLog.WriteEntry(evtLogSrc, string.Format("Run started {0}.", runStartTime.ToString()), EventLogEntryType.Information);
        string groupPrefix = "";
        int cycleOffset = -1;
        string accountsLstNm = "Accounts";
        string ratesLstNm = "Rates";
        string fixedConsumptionsLstNm = "Fixed Consumptions";
        string consumptionsLstNm = "Consumptions";
        string chargesLstNm = "Charges";
        string billingPeriodStr = "1m";
        bool isCycleOpen = false;
        bool? incremental = null;
        string lastRunFileName = "billease_last_run.log";
        DateTime cycleCalibrationDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, DateTimeKind.Local);

        Dictionary<string, List<KeyValuePair<string, string>>> listColumnsToCopy = new Dictionary<string, List<KeyValuePair<string, string>>>();
        listColumnsToCopy.Add("Account", new List<KeyValuePair<string, string>>());
        listColumnsToCopy.Add("Rate", new List<KeyValuePair<string, string>>());
        listColumnsToCopy.Add("Consumption", new List<KeyValuePair<string, string>>());
        var options = new OptionSet(){
                    {"p|prefix_of_group=", v => groupPrefix = v}
                    ,{"o|offset_of_cycle=", v => cycleOffset = int.Parse(v)}
                    ,{"b|billing_period=", v => billingPeriodStr = v}
                    ,{"d|cycle_calibration_date=", v => cycleCalibrationDate = DateTime.ParseExact(v,"yyyy-MM-dd",CultureInfo.InvariantCulture)}
                    ,{"a|accounts_list_name=", v => accountsLstNm = v}
                    ,{"r|rates_list_name=", v => ratesLstNm = v}
                    ,{"f|fixed_consumptions_list_name=", v => fixedConsumptionsLstNm = v}
                    ,{"c|consumptions_list_name=", v => consumptionsLstNm = v}
                    ,{"h|charges_list_name=", v => chargesLstNm = v}
                    ,{"A|account_columns_to_copy=", v => {
                      var src = v;
                      var dst = v;
                      if(v.Contains(":")){
                        src = v.Substring(0,v.IndexOf(":"));
                        dst = v.Substring(v.IndexOf(":")+1);
                      }
                      listColumnsToCopy["Account"].Add(new KeyValuePair<string, string>(src,dst));
                      }
                    }
                    ,{"R|rate_columns_to_copy=", v => {
                      var src = v;
                      var dst = v;
                      if(v.Contains(":")){
                        src = v.Substring(0,v.IndexOf(":"));
                        dst = v.Substring(v.IndexOf(":")+1);
                      }
                      listColumnsToCopy["Rate"].Add(new KeyValuePair<string, string>(src,dst));}
                    }
                    ,{"C|consumption_columns_to_copy=", v => {
                                           var src = v;
                      var dst = v;
                      if(v.Contains(":")){
                        src = v.Substring(0,v.IndexOf(":"));
                        dst = v.Substring(v.IndexOf(":")+1);
                      }
                      listColumnsToCopy["Consumption"].Add(new KeyValuePair<string, string>(src,dst));
                      }
                    }
                    ,{"O|is_cycle_open=", v=> isCycleOpen = Convert.ToBoolean(v)}
                    ,{"i|incremental=", v=> incremental = Convert.ToBoolean(v)}
                    ,{"l|last_run_log_file_name=", v=> lastRunFileName = v}
                };
        List<String> extraArgs = options.Parse(args);
        if (incremental == null)
        {
          incremental = isCycleOpen;
        }
        ServicePointManager.ServerCertificateValidationCallback = MyCertHandler;
        var cc = new ClientContext(extraArgs[0]);
        cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

        Match billingPeriodMatch = Regex.Match(billingPeriodStr, @"(\d)([mdy])");
        int billingPeriod = 0;
        string billingPeriodUOM = null;
        if (billingPeriodMatch.Success)
        {
          billingPeriod = Int32.Parse(billingPeriodMatch.Groups[1].Value);
          billingPeriodUOM = billingPeriodMatch.Groups[2].Value;
        }

        var billingCycleStart = DateTime.Now;
        TimeSpan s = billingCycleStart - cycleCalibrationDate;
        switch (billingPeriodUOM)
        {
          case "d":
            billingCycleStart = cycleCalibrationDate.AddDays(s.TotalDays - (s.TotalDays % billingPeriod) + cycleOffset * billingPeriod);
            break;
          case "m":
            int monthDiff = (billingCycleStart.Year - cycleCalibrationDate.Year) * 12 + billingCycleStart.Month - cycleCalibrationDate.Month;
            billingCycleStart = cycleCalibrationDate.AddMonths(monthDiff - (monthDiff % billingPeriod) + cycleOffset * billingPeriod);
            break;
          case "y":
            int yearDiff = (int)s.TotalDays / 365;
            billingCycleStart = cycleCalibrationDate.AddYears(yearDiff - (yearDiff % billingPeriod) + cycleOffset * billingPeriod);
            break;
        }
        DateTime nextBillingcycleStart = DateTime.Now;
        switch (billingPeriodUOM)
        {
          case "d":
            nextBillingcycleStart = billingCycleStart.AddDays(billingPeriod);
            break;
          case "m":
            nextBillingcycleStart = billingCycleStart.AddMonths(billingPeriod);
            break;
          case "y":
            nextBillingcycleStart = billingCycleStart.AddYears(billingPeriod);
            break;
        }
        TimeRange billingRange = new TimeRange(billingCycleStart, nextBillingcycleStart);

        // get last run timestamp
        var lastRunTs = new DateTime(1970, 01, 01);
        try
        {
          string[] lines = System.IO.File.ReadAllLines(lastRunFileName);
          DateTime lastRunCycle = DateTime.Parse(lines[0]);
          if (lastRunCycle.Date == billingCycleStart.Date)
          {
            lastRunTs = DateTime.Parse(lines[1]);
          }
        }
        catch
        {
        }


        // delete consumptions associated with deleted fixed consumptions
        var query = new CamlQuery();
        query.ViewXml = string.Format(@"<View><Query>
   <Where>
    <And>
      <Eq>
          <FieldRef Name='Cycle' />
          <Value Type='DateTime'>{0}</Value>
      </Eq>
      <And>
        <IsNull>
           <FieldRef Name='Fixed_x0020_Consumption_x0020_Re' />
        </IsNull>
        <Geq>
           <FieldRef Name='Fixed_x0020_Consumption_x0020_Re' LookupId='TRUE' />
           <Value Type=”Lookup”>0</Value>
        </Geq>
      </And>
    </And>
   </Where>
</Query></View>", billingCycleStart.ToString("yyyy-MM-dd"));
        var consumptionLst = cc.Web.Lists.GetByTitle(consumptionsLstNm);
        var consumptionFC = consumptionLst.Fields;
        var consumptionDeletionLIC = consumptionLst.GetItems(query);
        cc.Load(consumptionFC);
        cc.Load(consumptionDeletionLIC);
        cc.ExecuteQuery();
        foreach (var consumptionLI in consumptionDeletionLIC)
        {
          consumptionLI.DeleteObject();
          cc.ExecuteQuery();
        }

        // delete charges associated with deleted consumptions
        query = new CamlQuery();
        query.ViewXml = string.Format(@"<View><Query>
   <Where>
    <And>
      <Eq>
          <FieldRef Name='Cycle' />
          <Value Type='DateTime'>{0}</Value>
      </Eq>
      <And>
        <IsNull>
           <FieldRef Name='Consumption_x0020_Ref' />
        </IsNull>
        <Geq>
           <FieldRef Name='Consumption_x0020_Ref' LookupId='TRUE' />
           <Value Type=”Lookup”>0</Value>
        </Geq>
      </And>
    </And>
   </Where>
</Query></View>", billingCycleStart.ToString("yyyy-MM-dd"));
        var chgLst = cc.Web.Lists.GetByTitle(chargesLstNm);
        var chargesDeletionLIC = chgLst.GetItems(query);
        cc.Load(chargesDeletionLIC);
        cc.ExecuteQuery();
        foreach (var chargeLI in chargesDeletionLIC)
        {
          chargeLI.DeleteObject();
          cc.ExecuteQuery();
        }

        // populate or update consumptions from fixed consumptions
        query = new CamlQuery();
        query.ViewXml = string.Format(@"
<View><Query>
   <Where>
     <And>
        <Or>
          <IsNull>
             <FieldRef Name='Service_x0020_Start' />
          </IsNull>
          <Lt>
             <FieldRef Name='Service_x0020_Start' />
             <Value Type='DateTime'>{0}</Value>
          </Lt>
        </Or>
        <Or>
          <IsNull>
             <FieldRef Name='Service_x0020_End' />
          </IsNull>
          <Gt>
             <FieldRef Name='Service_x0020_End' />
             <Value Type='DateTime'>{1}</Value>
          </Gt>
        </Or>
     </And>
   </Where>
</Query></View>", nextBillingcycleStart.ToString("yyyy-MM-dd"), billingCycleStart.ToString("yyyy-MM-dd"));
        var fixedConsumptionsLst = cc.Web.Lists.GetByTitle(fixedConsumptionsLstNm);
        var fixedConsumptionLIC = fixedConsumptionsLst.GetItems(query);
        FieldCollection fixedConsumptionFC = fixedConsumptionsLst.Fields;
        cc.Load(fixedConsumptionFC);
        cc.Load(fixedConsumptionLIC);
        try
        {
          cc.ExecuteQuery();
          foreach (var fixedConsumptionLI in fixedConsumptionLIC)
          {
            // check if consumption exists for the current billing cycle
            var consumptionItemQuery = new CamlQuery();
            consumptionItemQuery.ViewXml = string.Format(@"
<View><Query>
   <Where>
    <And>
        <Eq>
            <FieldRef Name='Fixed_x0020_Consumption_x0020_Re' LookupId='TRUE' />
            <Value Type='Lookup'>{0}</Value>
        </Eq>
        <Eq>
            <FieldRef Name='Cycle' />
            <Value Type='DateTime'>{1}</Value>
        </Eq>
    </And>
   </Where>
</Query></View>", fixedConsumptionLI["ID"], billingCycleStart.ToString("yyyy-MM-dd"));
            var _consumptionLIC = consumptionLst.GetItems(consumptionItemQuery);
            cc.Load(_consumptionLIC);
            cc.ExecuteQuery();
            ListItem consumptionItem;
            if (_consumptionLIC.Count > 0)
            {
              if (isCycleOpen && ((DateTime)fixedConsumptionLI["Modified"]) < ((DateTime)_consumptionLIC[0]["Modified"]))
              {
                continue;
              }
              consumptionItem = _consumptionLIC[0];
            }
            else
            {
              ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
              consumptionItem = consumptionLst.AddItem(itemCreateInfo);
            }
            foreach (Field field in fixedConsumptionFC)
            {
              if (field.FromBaseType && field.InternalName != "Title")
              {
                continue;
              }

              if (consumptionFC.FirstOrDefault(f =>
                   (!f.FromBaseType || f.InternalName == "Title") && f.InternalName == field.InternalName
              ) == null)
              {
                continue;
              }
              consumptionItem[field.InternalName] = fixedConsumptionLI[field.InternalName];
            }

            // calculate proration
            try
            {
              if (fixedConsumptionLI["Prorated"] != null &&
                fixedConsumptionLI["Prorated"].ToString().Contains("Yes")
                )
              {
                DateTime serviceStart = DateTime.MinValue;
                DateTime serviceEnd = DateTime.MaxValue;
                if (fixedConsumptionLI["Service_x0020_Start"] != null)
                {
                  serviceStart = ((DateTime)fixedConsumptionLI["Service_x0020_Start"]).ToLocalTime();
                }
                if (fixedConsumptionLI["Service_x0020_End"] != null)
                {
                  serviceEnd = ((DateTime)fixedConsumptionLI["Service_x0020_End"]).ToLocalTime();
                }
                TimeRange serviceRange = new TimeRange(serviceStart, serviceEnd);
                ITimeRange overlap = serviceRange.GetIntersection(billingRange);
                double portion = overlap.Duration.TotalDays / billingRange.Duration.TotalDays;
                if (fixedConsumptionLI["Quantity"] != null)
                {
                  var proratedQty = ((double)fixedConsumptionLI["Quantity"]) * portion;
                  if (fixedConsumptionLI["Prorated"].ToString() != "Yes")
                  {
                    proratedQty = Convert.ToInt32(proratedQty);
                  }
                  consumptionItem["Quantity"] = proratedQty;
                }
                if (fixedConsumptionLI["Amount"] != null)
                {
                  consumptionItem["Amount"] = ((double)fixedConsumptionLI["Amount"]) * portion;
                }
              }
            }
            catch
            {
              EventLog.WriteEntry(evtLogSrc, string.Format("Error calculating proration for fixed consumption with ID={0}", fixedConsumptionLI["ID"]), EventLogEntryType.Error);
            }

            consumptionItem["Cycle"] = billingCycleStart;
            FieldLookupValue lookup = new FieldLookupValue();
            lookup.LookupId = (int)fixedConsumptionLI["ID"];
            consumptionItem["Fixed_x0020_Consumption_x0020_Re"] = lookup;
            consumptionItem.Update();
            cc.ExecuteQuery();
          }
        }
        catch (Exception ex)
        {
          EventLog.WriteEntry(evtLogSrc, string.Format("Error creating consumption from fixed consumption.\n{0}", ex.ToString()), EventLogEntryType.Error);
        }


        // populate or update charge from consumption
        query = new CamlQuery();
        string viewXml = null;
        if (incremental == false)
        {
          viewXml = string.Format(@"
<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='Cycle' />
         <Value Type='DateTime'>{0}</Value>
      </Eq>
   </Where>
</Query></View>", billingCycleStart.ToString("yyyy-MM-dd"));
        }
        else
        {
          viewXml = string.Format(@"
<View><Query>
   <Where>
    <And>
      <Eq>
         <FieldRef Name='Cycle' />
         <Value Type='DateTime'>{0}</Value>
      </Eq>
      <Gt>
         <FieldRef Name='Modified' />
         <Value IncludeTimeValue='true' Type='DateTime'>{1}</Value>
      </Gt>
    </And>
   </Where>
</Query></View>", billingCycleStart.ToString("yyyy-MM-dd"), lastRunTs.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
        query.ViewXml = viewXml;

        var consumptionLIC = consumptionLst.GetItems(query);
        cc.Load(consumptionLIC, items => items.IncludeWithDefaultProperties(
            item => item["HasUniqueRoleAssignments"]
            ));
        cc.Load(cc.Web.RoleDefinitions);
        var gc = cc.Web.SiteGroups;
        cc.Load(gc);
        cc.ExecuteQuery();
        var restReadRD = cc.Web.RoleDefinitions.GetByName("Restricted Read");
        var readRD = cc.Web.RoleDefinitions.GetByName("Read");

        foreach (var consumptionLI in consumptionLIC)
        {
          ListItem chgItem;
          // check if charges exists
          var chgItemQuery = new CamlQuery();
          chgItemQuery.ViewXml = string.Format(@"
<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='Consumption_x0020_Ref' LookupId='TRUE' />
         <Value Type='Lookup'>{0}</Value>
      </Eq>
   </Where>
</Query></View>", consumptionLI["ID"]);
          var chgLIC = chgLst.GetItems(chgItemQuery);
          cc.Load(chgLIC);
          cc.ExecuteQuery();
          if (chgLIC.Count > 0)
          {
            chgItem = chgLIC[0];
            if (isCycleOpen && ((DateTime)consumptionLI["Modified"]) < ((DateTime)chgItem["Modified"]))
            {
              continue;
            }
            chgItem.ResetRoleInheritance();
            cc.ExecuteQuery();
          }
          else
          {
            // create new charges item
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            chgItem = chgLst.AddItem(itemCreateInfo);
          }

          // get org item
          var orgItemQuery = new CamlQuery();
          orgItemQuery.ViewXml = string.Format(@"
<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='ID' />
         <Value Type='Counter'>{0}</Value>
      </Eq>
   </Where>
</Query></View>", ((FieldLookupValue)consumptionLI["Account"]).LookupId);
          var orgLst = cc.Web.Lists.GetByTitle(accountsLstNm);
          var orgLIC = orgLst.GetItems(orgItemQuery);
          cc.Load(orgLIC);
          // get rate item
          var rateItemQuery = new CamlQuery();
          rateItemQuery.ViewXml = string.Format(@"<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='ID' />
         <Value Type='Counter'>{0}</Value>
      </Eq>
   </Where>
</Query></View>", ((FieldLookupValue)consumptionLI["Rate"]).LookupId);
          var rateLst = cc.Web.Lists.GetByTitle(ratesLstNm);
          var rateLIC = rateLst.GetItems(rateItemQuery);
          cc.Load(rateLIC);
          cc.ExecuteQuery();
          var orgItem = orgLIC.First();
          var rateItem = rateLIC.First();

          chgItem["Account"] = orgItem["Title"];
          chgItem["Title"] = consumptionLI["Title"];
          chgItem["Cycle"] = consumptionLI["Cycle"];
          chgItem["Unit_x0020_Price"] = rateItem["Unit_x0020_Price"];
          chgItem["Denominator"] = rateItem["Denominator"];
          chgItem["UOM"] = rateItem["UOM"];
          chgItem["Quantity"] = consumptionLI["Quantity"];
          FieldLookupValue lookup = new FieldLookupValue();
          lookup.LookupId = (int)consumptionLI["ID"];
          chgItem["Consumption_x0020_Ref"] = lookup;

          // the order of list enumeration is important
          foreach (var lstNm in new string[] { "Account", "Rate", "Consumption" })
          {
            KeyValuePair<string, List<KeyValuePair<string, string>>> listColumnToCopy = new KeyValuePair<string, List<KeyValuePair<string, string>>>(lstNm, listColumnsToCopy[lstNm]);
            if (listColumnToCopy.Value.Count <= 0)
            {
              continue;
            }
            ListItem item = null;
            switch (listColumnToCopy.Key)
            {
              case "Account":
                item = orgItem;
                break;
              case "Rate":
                item = rateItem;
                break;
              case "Consumption":
                item = consumptionLI;
                break;
            }
            foreach (var columnNVP in listColumnToCopy.Value)
            {
              try
              {
                if (item[columnNVP.Key] != null)
                {
                  chgItem[columnNVP.Value] = item[columnNVP.Key];
                }
              }
              catch
              {
                EventLog.WriteEntry(evtLogSrc, string.Format(@"Cannot copy column {0} in list {1} for consumption ID={2}", columnNVP.Key, listColumnToCopy.Key, consumptionLI["ID"]), EventLogEntryType.Error);
              }
            }

          }

          if (consumptionLI["Amount"] != null)
            chgItem["Amount"] = consumptionLI["Amount"];
          else if (consumptionLI["Quantity"] != null
              && rateItem["Denominator"] != null
              && rateItem["Unit_x0020_Price"] != null
              && rateItem["Denominator"] != null
              && (double)rateItem["Denominator"] > 0)
          {
            var normalizedQty = (double)consumptionLI["Quantity"] / (double)rateItem["Denominator"];
            try
            {

              if (rateItem["Round_x0020_Up"] != null && rateItem["Round_x0020_Up"] as Nullable<bool> == true)
              {
                normalizedQty = Math.Ceiling((double)consumptionLI["Quantity"] / (double)rateItem["Denominator"]);
              }
            }
            catch
            {
              EventLog.WriteEntry(evtLogSrc, string.Format("Error calculate round up for rate ID={0}, consumption ID={1}", rateItem["ID"], consumptionLI["ID"]), EventLogEntryType.Error);
            }
            chgItem["Amount"] = (double)rateItem["Unit_x0020_Price"] * normalizedQty;
          }
          if (chgItem.FieldValues.ContainsKey("Amount"))
          {
            chgItem.Update();
            cc.ExecuteQuery();
          }
          else
          {
            EventLog.WriteEntry(evtLogSrc, "Cannot calculate amount for consumption item " + consumptionLI["ID"], EventLogEntryType.Error);
          }

          if (!isCycleOpen)
          {
            if (!consumptionLI.HasUniqueRoleAssignments)
            {
              consumptionLI.BreakRoleInheritance(true, false);
              cc.ExecuteQuery();
            }
            cc.Load(consumptionLI.RoleAssignments, items => items.Include(
                ra => ra.RoleDefinitionBindings.Include(
                    rdb => rdb.Name
                    )
                ));
            cc.ExecuteQuery();

            foreach (var ra in consumptionLI.RoleAssignments)
            {
              bool addRead = false;
              foreach (var rdbo in ra.RoleDefinitionBindings)
              {
                if (rdbo.Name == "Contribute")
                {
                  ra.RoleDefinitionBindings.Remove(rdbo);
                  addRead = true;
                }
              }
              if (addRead) ra.RoleDefinitionBindings.Add(readRD);
              ra.Update();
            }
            cc.ExecuteQuery();
          }
        }

        var chargesLst = cc.Web.Lists.GetByTitle("Charges");
        var chargesLIC = chargesLst.GetItems(query);
        cc.Load(chargesLIC, items => items.Include(
            item => item["Account"]
            , item => item["HasUniqueRoleAssignments"]
            ));
        cc.ExecuteQuery();
        foreach (var chargeLI in chargesLIC)
        {
          if (!chargeLI.HasUniqueRoleAssignments)
          {
            chargeLI.BreakRoleInheritance(true, false);
          }
          foreach (var g in gc)
          {
            if (g.LoginName == (groupPrefix + chargeLI["Account"].ToString()))
            {
              var rdb = new RoleDefinitionBindingCollection(cc);
              rdb.Add(restReadRD);
              chargeLI.RoleAssignments.Add(g, rdb);
              break;
            }
          }
          cc.ExecuteQuery();
        }
        var runEndTime = DateTime.Now;
        EventLog.WriteEntry(evtLogSrc, string.Format("Run ended {0}, lasting {1} minutes.", runEndTime.ToString(), (runEndTime - runStartTime).TotalMinutes.ToString("0.00")), EventLogEntryType.Information);
        System.IO.File.Delete(lastRunFileName);
        System.IO.File.WriteAllLines(lastRunFileName, new string[] { billingCycleStart.ToShortDateString(), runStartTime.ToString() });
      }
      catch (Exception ex)
      {
        EventLog.WriteEntry(evtLogSrc, ex.ToString(), EventLogEntryType.Error);
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
