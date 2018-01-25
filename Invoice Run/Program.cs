﻿using System;
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

namespace Invoice_Run
{
	class Program
	{
		const string evtLogSrc = "Service Billing Invoice Run";
		static void Main(string[] args)
		{
			if (!EventLog.SourceExists(evtLogSrc))
			{
				EventLog.CreateEventSource(evtLogSrc, "Application");
			}

			try
			{
				string groupPrefix = "";
				int cycleMonthOffset = -1;
				string accountsLstNm = "Accounts";
				string ratesLstNm = "Rates";
				string fixedConsumptionsLstNm = "Fixed Consumptions";
				string consumptionsLstNm = "Consumptions";
				string chargesLstNm = "Charges";
				string billingPeriodStr = "1m";

				Dictionary<string, List<string>> listColumnsToCopy = new Dictionary<string, List<string>>();
				listColumnsToCopy.Add("Account", new List<string>());
				listColumnsToCopy.Add("Rate", new List<string>());
				listColumnsToCopy.Add("Consumption", new List<string>());
				var options = new OptionSet(){
                    {"p|prefix_of_group=", v => groupPrefix = v}
                    ,{"o|offset_of_cycle_month=", v => cycleMonthOffset = int.Parse(v)}
                    ,{"b|billing_period=", v => billingPeriodStr = v}
                    ,{"a|accounts_list_name=", v => accountsLstNm = v}
                    ,{"r|rates_list_name=", v => ratesLstNm = v}
                    ,{"f|fixed_consumptions_list_name=", v => fixedConsumptionsLstNm = v}
                    ,{"c|consumptions_list_name=", v => consumptionsLstNm = v}
                    ,{"h|charges_list_name=", v => chargesLstNm = v}
                    ,{"n|account_columns_to_copy=", v => listColumnsToCopy["Account"].Add(v)}
                    ,{"t|rate_columns_to_copy=", v => listColumnsToCopy["Rate"].Add(v)}
                    ,{"s|consumption_columns_to_copy=", v => listColumnsToCopy["Consumption"].Add(v)}
                };
				List<String> extraArgs = options.Parse(args);
				ServicePointManager.ServerCertificateValidationCallback = MyCertHandler;
				var cc = new ClientContext(extraArgs[0]);
				cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

				var billingCycleStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
				billingCycleStart = billingCycleStart.AddMonths(cycleMonthOffset);
				Match billingPeriodMatch = Regex.Match(billingPeriodStr, @"(\d)([mdy])");
				int billingPeriod = 0;
				string billingPeriodUOM = null;
				if (billingPeriodMatch.Success)
				{
					billingPeriod = Int32.Parse(billingPeriodMatch.Groups[1].Value);
					billingPeriodUOM = billingPeriodMatch.Groups[2].Value;
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
				var query = new CamlQuery();
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
				var consumptionLst = cc.Web.Lists.GetByTitle(consumptionsLstNm);
				var consumptionFC = consumptionLst.Fields;
				cc.Load(consumptionFC);
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
						if (_consumptionLIC.Count > 0)
						{
							continue;
						}

						ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
						var newConsumptionItem = consumptionLst.AddItem(itemCreateInfo);
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
							newConsumptionItem[field.InternalName] = fixedConsumptionLI[field.InternalName];
						}
						newConsumptionItem["Cycle"] = billingCycleStart;
						FieldLookupValue lookup = new FieldLookupValue();
						lookup.LookupId = (int)fixedConsumptionLI["ID"];
						newConsumptionItem["Fixed_x0020_Consumption_x0020_Re"] = lookup;
						newConsumptionItem.Update();
						cc.ExecuteQuery();
					}
				}
				catch (Exception ex)
				{
					EventLog.WriteEntry(evtLogSrc, string.Format("Error creating consumption from fixed consumption.\n{0}", ex.ToString()), EventLogEntryType.Error);
				}

				query = new CamlQuery();
				query.ViewXml = string.Format(@"
<View><Query>
   <Where>
      <Eq>
         <FieldRef Name='Cycle' />
         <Value Type='DateTime'>{0}</Value>
      </Eq>
   </Where>
</Query></View>", billingCycleStart.ToString("yyyy-MM-dd"));

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
				var chgLst = cc.Web.Lists.GetByTitle(chargesLstNm);

				foreach (var consumptionLI in consumptionLIC)
				{
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
						continue;
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

					// create new charges item
					ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
					var newChgItem = chgLst.AddItem(itemCreateInfo);
					newChgItem["Account"] = orgItem["Title"];
					newChgItem["Title"] = consumptionLI["Title"];
					newChgItem["Cycle"] = consumptionLI["Cycle"];
					newChgItem["Unit_x0020_Price"] = rateItem["Unit_x0020_Price"];
					newChgItem["Denominator"] = rateItem["Denominator"];
					newChgItem["UOM"] = rateItem["UOM"];
					newChgItem["Quantity"] = consumptionLI["Quantity"];
					FieldLookupValue lookup = new FieldLookupValue();
					lookup.LookupId = (int)consumptionLI["ID"];
					newChgItem["Consumption_x0020_Ref"] = lookup;

					foreach (KeyValuePair<string, List<string>> listColumnToCopy in listColumnsToCopy)
					{
						if (listColumnToCopy.Value.Count <= 0)
						{
							continue;
						}
						ListItem item = null;
						switch (listColumnToCopy.Key)
						{
							case "Rate":
								item = rateItem;
								break;
							case "Account":
								item = orgItem;
								break;
							case "Consumption":
								item = consumptionLI;
								break;
						}
						foreach (var columnNm in listColumnToCopy.Value)
						{
							try
							{
								newChgItem[columnNm] = item[columnNm];
							}
							catch (Exception)
							{
								EventLog.WriteEntry(evtLogSrc, string.Format(@"Cannot copy column {0} in list {1}", columnNm, listColumnToCopy.Key), EventLogEntryType.Error);
							}
						}

					}

					if (consumptionLI["Amount"] != null)
						newChgItem["Amount"] = consumptionLI["Amount"];
					else if (consumptionLI["Quantity"] != null
							&& rateItem["Denominator"] != null
							&& rateItem["Unit_x0020_Price"] != null
							&& rateItem["Denominator"] != null
							&& (double)rateItem["Denominator"] > 0)
						newChgItem["Amount"] = (double)rateItem["Unit_x0020_Price"] * Math.Ceiling((double)consumptionLI["Quantity"] / (double)rateItem["Denominator"]);
					if (newChgItem.FieldValues.ContainsKey("Amount"))
					{
						newChgItem.Update();
						cc.ExecuteQuery();
					}
					else
					{
						EventLog.WriteEntry(evtLogSrc, "Cannot calculate amount for consumption item " + consumptionLI["ID"], EventLogEntryType.Error);
					}

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
