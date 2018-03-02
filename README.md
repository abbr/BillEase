BillEase
==========================

A service billing SharePoint web application designed for the purpose of intra-organization cost allocation supporting recurring charges.

## Background
In a large organization, a support division often provides ongoing services to other divisions. For example, IT department maintains a web site for marketing department. To better keep track of cost allocation, it is often necessary to price the services and charge the client divisions based on certain consumption metrics, similar to what utility companies do to you. It is more so if the support division adopts cost-recovery model. Unlike the case of utility companies, however, implicit trust exists among divisions within an organization. Therefore the requirements of invoicing, payment and collection can be greatly simplified or even eliminated. What's left over are, in essence:

  * Defining services and rates
  * Periodical (monthly, for example) metering
  * Calculating charges based on rates and meter readings at the end of each billing cycle
  * Providing a customer portal for clients to retrieve their consumption and charges. 

*BillEase* is a service billing web application to support above trimmed down functions. It is not, however, intended to replace a corporate accounting system that tracks the flow of funds from one cost center to another.

## Features

  * Support both recurring and one-time fixed rate charges and credits
  * Support proration of quantity for fixed quantity recurring charges in the billing cycles covering the begining or end of service period.
  * Allow overriding calculated line item amount
  * Prevent altering charge and consumption records of closed billing cycles
  * Features inherited from SharePoint
     * Bulk importing and updating
     * Versioning
     * Single-Sign-On with enterprise-wide identity management system, assuming the system is Active Directory
     * Permission at a granularity down to item level
     * Auditing information such as creator, last modifier, creation time and last modified time
     * Recycle Bin
     * Export to Excel

## Overview
*BillEase* consists of five SharePoint custom lists packaged into a SharePoint site template solution file *Service Billing.wsp* and a console application called *Invoice Run.exe*. The custom lists are:
  * *Accounts* - contains client information
  * *Rates* - defines services and corresponding rates
  * *Consumptions* - is used to upload meter readings
  * *Charges* - auto-populated line items with amount computed from rates and consumption. 
  * *Fixed Consumptions* - contains pre-defined fixed-quantity consumptions to be posted to *Consumptions* list periodically

*Invoice Run.exe* is intended to be put into a scheduled task to run at the end of each billing cycle to create charge line items.

*BillEase* is intended to be accessed by four types of users with following permissions:

1. Administrator (owner) from service provider - full control to entire site.
2. Contributor (member) from service provider - full read access to entire site plus write access to *Consumptions* list. This type should contain service delivery unit personnel.
3. Visitor - full read access to entire site. This can be accountants from service provider or external auditors, for example.
4. Client - read access to security trimmed *Charges* list.

*BillEase* requires following manual processes performed by the service provider:

1. One-time activities - performed first time as initial setup and update only when information changes thereafter by administrator.
   1. Populate *Accounts* and *Rates*. 
   2. For each *Account*, create a SharePoint group named after it. A prefix is allowed. For example, if the account is called *Marketing*, the corresponding SharePoint group should be called *Billing Group - Marketing*, assuming prefix is *Billing Group -*. The prefix has to be consistent. Change the group name whenever account name changes. Don't grant the group any permission to the site.
   3. Add client users who are allowed to see charges to an account to the corresponding groups.
   4. Setup and adjust permissions as needed. For example, when creating the SharePoint site from solution, by default site contributors have write permission to all lists. To match the access permission described above, you may want to change the default permission to read-only, then grant contributors write access to *Consumptions* list by breaking the permission inheritance.
   5. (Optional) Populate pre-defined fixed-quantity consumptions in *Fixed Consumptions* list
2. Recurring activities:
   1. Near the end of each billing cycle, stage meter readings except for *Fixed Consumptions* into an Excel spreadsheet with column order matching the *Consumptions* list view.
   2. Select the range of data in Excel and press Ctrl-C to copy. Open the *Datasheet View* of *Consumptions* list in Internet Explorer. click the second column (by default *Title*) of the last empty row in the *Datasheet View* marked by asterisk, and press Ctrl-V to paste the data range to the list. This completes the bulk loading process. Data uploaded can be modified as long as the billing cycle is not closed.

Once above activities are performed, rest processes are handled automatically by *Invoice Run.exe*. Clients can see their charge line items in *Charges* list. They can export the list to Excel for further analysis.

## Components

### SharePoint Custom Lists
*BillEase* depends on list and column names described below to function. Extending a list is allowed as long as these columns are not deleted. Renaming these columns are allowed, however. To rename a list, make sure to supply the new name to *Invoice Run.exe* documented in section [Console Application](#console-application) below.

#### Accounts
*Accounts* contain client information. Only *Title* column is mandatory. Changing the name of an account is allowed. However, the value of *Account* column in *Charges* list described below is copied from, not referencing to, the *Title* column of *Accounts* list, so the account name change will not propagate to *Charges* list.
Deleting an account is disallowed unless all consumption items associated with the account are deleted.

#### Rates
*Rates* define billable line item types and corresponding rates. It has following columns at minimum:
  * Title - name of the billable line item type. This could be the service name if the line item type is service specific.
  * Unit Price
  * UOM - unit of measure
  * Denominator - a number combined with UOM to form the denominator of the rate. Denominator is used for round-up calculation if enabled by *Round Up* column described next. For example, let's say the service being provided is data storage, and the price is $10 per 5GB per billing cycle. In this case Unit Price is $10, UOM is GB, and denominator is 5. When priced this way, consumption is rounded-up at increment of 5GB. For instance, 6GB costs $20, as opposed to $12 had the round up been disabled.
  * Round Up - Whether or not round up consumed quantity to next increment of denominator when calculating amount in charges. Default to yes.

Changing *Unit Price* or *Denominator* only affects future charge calculations. Deleting a rate entry is disallowed unless all consumption items associated with the rate are deleted.

#### Consumptions
*Consumptions* list is used to upload meter readings. It has following columns at minimum:
  * Title - consumption title. This becomes the charge line item description.
  * Account - a reference to the account
  * Rate - a reference to the rate
  * Quantity - consumption data. UOM should match that of rate referenced
  * Cycle - the  start of billing cycle. The first day of current month is populated by default.
  * Amount - used to override calculated amount. This column is useful to post one-off type of charges or credits. When this column is populated, *Quantity* doesn't need to be populated. Even if *Quantity* is populated, the quantity will not be used.
  * Fixed Consumption Ref - if the consumption item is auto-populated from *Fixed Consumptions* by *Invoice Run.exe*, this hidden field contains a reference to the corrseponding fixed consumption item.

Consumption items are modifiable by users with *Contribute* permission of the list prior to the closing date of billing cycle and read-only thereafter. Consumption items are also modifiable by administrators any time. Deleting a consumption item of a closed billing cycle is disallowed unless the corresponding charge item is deleted.
 
#### Charges
Items in *Charges* list are created by *Invoice Run.exe*. There is a one-to-one mapping between *Consumptions* and *Charges*. *Charges* contain following columns at minimum:
  * Title - copied from  *Consumptions*
  * Account - copied from *Consumptions*
  * Cycle - copied from *Consumptions*
  * Unit Price - copied from *Rates* referenced by the corresponding *Consumptions* item
  * Denominator - copied from *Rates* referenced by the corresponding *Consumptions* item
  * UOM - copied from *Rates* referenced by the corresponding *Consumptions* item
  * Quantity - copied from *Consumptions*
  * Amount - either copied from *Consumptions* or, in absence of value, calculated using formula *Unit Price\*Ceiling(Quantity/Denominator)*
  * Consumption Ref  - a hidden field referencing to the corresponding *Consumptions* item

Notice that except for the hidden *Consumption Ref* column, all columns are copied from, rather than referencing to other lists. This *de-normalization* process prevents historical billing records from altering by factors such as account re-naming or price adjustment, resulting in improved accountability.

By the same record-preserving principle, charge line items should be made read-only, except for site collection administrators who have full access regardless of permissions. When a charge item is created, the permission of the item is broken from inheritance. Users who have read permissions defined in the *Charges* list at the time of broken can still read the item. In addition, users who belong to the *"&lt;prefix&gt;&lt;account&gt;"* group are also granted read-only access. This makes the list security-trimmed and suitable to be exposed as a portal page to clients who can only see the charges applied to their account.

#### Fixed Consumptions
*Fixed Consumptions* list is a worksheet used to define consumptions of pre-determined quantities in a set-it-and-forget-it manner. The columns in this list generally match *Consumptions* list except for following extra columns
* Service Start - optional start date of the service
* Service End - optional end date of the service
* Prorated - whether or not to prorate Quantity and Amount during the billing cycles that cover the start or end of service period. Allowed values are
  * No
  * Yes
  * Yes and round quantity to integer
  
  Default to *No*.

When *Invoice Run.exe* is executed, a new consumption item is generated for every fixed consumption item with service period  \[<*Service Start*>, <*Service End*>) overlapping the billing cycle by copying columns exist in both lists. Note *Service Start* is inclusive and *Service End* is exclusive. Missing *Service Start* implies a distant past; missing *Service End* implies a distant future. 

*Invoice Run.exe* will also populate following *Consumptions* list columns that *Fixed Consumptions* shouldn't contain:

* *Cycle* is set to billing cycle start date. 
* *Fixed Consumption Ref* is set to a reference to the fixed consumption item. *Invoice Run.exe* relies on this column to avoid generating multiple consumption items in the same billing cycle in case *Invoice Run.exe* has to be executed repetitively.

Item level proration is supported through the *Prorated* list column. If this field is set to *Yes* or *Yes and round quantity to integer*, then the Quantity and Amount fields are prorated by day in the billing cycles that cover the service start or end date. For example, if the service start date is 2018-02-01 and service end date is empty, assuming billing period is quarterly starting January 1 annually, then for the billing cycle 2018-01-01 to 2018-03-31, Quantity and Amount fields are adjusted by a proration factor of 59/90≈0.66 when posting to the *Consumptions* list. Furthermore, if the field is set to *Yes and round quantity to integer*, then prorated *Quantity* field is rounded to integer. This allows *Quantity* field to be set correctly to accomodate the fact that months have different number of days. In the above example, if *Quantity* field is 3 (reasonable because rate can be priced as per month and there are 3 months in a quarter), then the rounded prorated quantity will be 2, as opposed to 1.966666666666667 had rounding not been performed.

### Console Application
The gem of *BillEase* is the console application *Invoice Run.exe*. It provides automation and turns the five SharePoint lists into a workable solution. Without it the SharePoint lists are merely data repository. *Invoice Run.exe* is intended to be launched by a scheduled task at the close of each billing cycle (by default first day of each month). For testing purpose it can also be launched manually and repetitively. When invoked, *Invoice Run.exe* performs following tasks:

1. For each fixed consumption item with service period overlapping the affected billing period, create a consumption item if not already exists.
2. For each consumption item in affected billing cycle, create a charge item if not already exists. The values of charge item are copied or calculated using data directly or indirectly obtained from consumption item as described in [Charges](#charges) list above.
3. Break the permission inheritance of each consumption item in affected billing cycle if not already done so. Then convert all *Contribute* permissions to *Read*.
4. Break the permission inheritance of each charge item created in affected billing cycle if not already done so. 
5. For each charge item in affected billing cycle, grant group *"&lt;prefix&gt;&lt;account&gt;"* read-only access if not already done so.

*Invoice Run.exe* expects following call syntax:
```
"Invoice Run.exe" [options] <URL>
where <URL> points to the site holding the five lists and [options] are
-p|--prefix_of_group=<string>
	Prefix of the account groups. The prefix is useful to 
    prevent group name conflicts with other groups defined in same site collection
-b|--billing_period=<string>
	Billing period in the form <Integer><UOM>. Default to 1m. Allowed UOMs are 
	<d|m|y> for day, month and year respectively. For example, 1m for 1 month; 
	14d for 2 weeks; 3m for a quarter.
-C|--cycle_calibration_date=<date>
	Billing cycle calibration date in the format YYYY-MM-DD. This is a past date 
	that is known to be the start of a billing cycle. Default value is the first day
	of the current month the program is running.
-o|--offset_of_cycle=<number>
	Offset of billing cycle. Default to -1. For example, 
	if billing cycle starts on the first day of each month and Invoice Run 
	is launched at 12:01AM on the first day of each month, the default offset 
	of -1 is needed for calculation be performed on last month's data.
-a|--accounts_list_name=<string>
	Name of accounts list if renamed.
-r|--rates_list_name=<string>
	Name of rates list if renamed.
-f|--fixed_consumptions_list_name=<string>
	Name of fixed consumptions list if renamed.
-c|--consumptions_list_name=<string>
	Name of consumptions list if renamed.
-h|--charges_list_name=<string>
	Name of charges list if renamed.
-n|--account_columns_to_copy=<string>
	Name of custom column in accounts list to copy over to charges list in the format 
	<source_name>[:<target_name>]. Target name can be omitted if same as source name.
	The column name is not display name, but rather internal field name which you can find in the 
	Field url query parameter when opening the column definition in list settings.
	Multiple columns can be defined by adding this option multiple times. The column type
	in both accounts and charges lists must match. If a column is also copied from rate or
	consumption list, the precedence of overriding in descending order is: consumption, rate, 
	account.
-t|--rate_columns_to_copy=<string>
	Name of custom column in rates list to copy over to charges list in the format <source_name>
	[:<target_name>]. Target name can be omitted if same as source name. 
	The column name is not display name, but rather internal field name which you can find in the 
	Field url query parameter when opening the column definition in list settings.
	Multiple columns can be	defined by adding this option multiple times. The column type	in both
	rates and charges lists	must match. If a column is also copied from account or consumption list, 
	the precedence of overriding in descending order is: consumption, rate,	account.
-s|--consumption_columns_to_copy=<string>
	Name of custom column in consumptions list to copy over to charges list in the format 
	<source_name>[:<target_name>]. Target name can be omitted if same as source name. 
	The column name is not display name, but rather internal field name which you can find in the 
	Field url query parameter when opening the column definition in list settings.
	Multiple columns can be defined by adding this option multiple times. The column type
	in both consumptions and charges lists must match. If a column is also copied from account or
	rate list, the precedence of overriding in descending order is: consumption, rate, account.
-O|--is_cycle_open=<true|false>
	Whether the billing cycle under operation is open or not. If not, then all 
	consumptions that have been posted to charges are frozen from updates by contributors. 
	This option tentatively has no effect. Default to false.
i|--incremental=<true|false>
	Whether or not update charges incrementally. If incremental, only consumptions modified 
	since last run will be fetched. By default, incremental is true if billing cycle 
	is open and false if closed.
-l|--last_run_log_file_name=<string>
	Name of last run log file name. Default to billease_last_run.log. If there are multiple 
	scheduled tasks created to run the console application and all scheduled tasks are set 
	to the same working directory, then each should have its own last run log file by 
	supplying a different name for this option.

Examples:
"Invoice Run.exe" -p "Billing Group - " https://mycorp.com/service/billing
  Set the prefix of all SharePoint groups used to grant clients accessing 
  the Charges table to "Billing Group - ". Client users from account 
  Marketing, for example, should be placed in a SharePoint group 
  called "Billing Group - Marketing" 
"Invoice Run.exe" -o 0 https://mycorp.com/service/billing
  Set the offset of billing cycle month adjustment to 0. This is needed if billing 
  cycle starts on the first day of each month and Invoice Run is launched at
  11:50PM on the last day of each month, for example.
"Invoice Run.exe" -c Comments -c ServiceDate:Service_x0020_Date https://mycorp.com/service/billing
  When creating charges items, copy Comments and ServiceDate columns in consumptons 
  list over to Comments and Service_x0020_Date columns respectively in charges list.
```
## System Requirements and Access Privileges
* Site collection administrator level of access to any edition of SharePoint 2010 or 2013 with integrated Windows authentication.
* Local administrator access to a server with .Net Framework 4 installed to run scheduled tasks. The server doesn't need to be the host of the SharePoint site. Windows Server 2008 R2 has been tested working.
* Optionally Git client to download package
* Optionally Visual Studio 2017 if you want to compile or change source code of  *Invoice Run.exe*

## Installation
1. Use *Download Zip* button to download the latest version, or use Git client to clone the git repo. 
2.  Upload file *Service Billing.wsp* to SharePoint site collection solution gallery and activate it. 
3. Create a site using *Service Billing* site template contained in *Service Billing.wsp*. You may need to active certain site collection features first.
4. Define permissions of each list appropriately. Users from service provider organization, depending on job roles, should have read-only permission to *Charges* and read-write permission to *Accounts*, *Rates* and *Consumptions*. Don't grant any permission to clients at list level.
5. Follow manual processes in [Overview](#overview) section to populate lists and SharePoint groups. Create some fake data in *Consumptions* list in order to verify the function.
6. Copy all files under */Invoice Run/bin/Debug* to a server where *Invoice Run* scheduled task will be created. The server must have .Net Framework 4 installed. 
7. Manually run *Invoice Run.exe* on the server with URL of the site created in Step 3. above and optional arguments documented in the [Console Application](#console-application) section above. The Windows log in account should be a site collection administrator as well as a local server administrator. If you run from a desktop version of Windows such as Vista with UAC, you have to run *Invoice Run.exe* from a DOS prompt started with "Run as administrator". If the run is successful, you should see new items created in *Charges* list with unique permissions. If the run fails, errors are output to both console and Windows event log.
8. Create a scheduled task to run *Invoice Run.exe* periodically. The account used to run the scheduled task should have adequate privilege to modify list items and permissions. Make the account a site collection administrator is recommended.

## License

The MIT License (MIT)

Copyright (c) 2014-present @abbr

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
