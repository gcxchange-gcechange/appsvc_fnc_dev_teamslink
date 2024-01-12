using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace appsvc_fnc_dev_teamslink
{
    public static class GetTeamsLink
    {
        // Runs at 06:00 on Sunday
        [FunctionName("GetTeamsLink")]
        public static async Task Run([TimerTrigger("0 0 6 * * 0")] TimerInfo myTimer, ILogger log)
        {
            log.LogInformation("GetTeamsLink received a request.");

            IConfiguration config = new ConfigurationBuilder()
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            var exceptionGroupsArray = config["exceptionGroupsArray"];
            var siteId = config["siteId"];
            var listId = config["listId"];
            var tenantid = config["tenantid"];

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            List<CreateItem> CreateList = new();
            List<Microsoft.Graph.Models.Group> groups = new();
            List<ListItem> items = new();
            List<ListItem> UpdateList = new();

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Get items from TeamsLink list                                                                          //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            var listitems = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=TeamsID,Teamslink)" };
                requestConfiguration.QueryParameters.Top = 10;  // 999
            });

            items.AddRange(listitems.Value);

            // fetch next page(s)
            while (listitems.OdataNextLink != null)
            {
                var nextPageRequestInformation = new RequestInformation
                {
                    HttpMethod = Method.GET,
                    UrlTemplate = listitems.OdataNextLink
                };

                listitems = await graphClient.RequestAdapter.SendAsync(nextPageRequestInformation, (parseNode) => new ListItemCollectionResponse());
                items.AddRange(listitems.Value);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Get groups from tenant                                                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            var listgroups = await graphClient.Groups.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id,resourceProvisioningOptions" };
                requestConfiguration.QueryParameters.Top = 999;
            });

            groups.AddRange(listgroups.Value);

            // fetch next page(s)
            while (listgroups.OdataNextLink != null)
            {
                var nextPageRequestInformation = new RequestInformation
                {
                    HttpMethod = Method.GET,
                    UrlTemplate = listgroups.OdataNextLink
                };

                listgroups = await graphClient.RequestAdapter.SendAsync(nextPageRequestInformation, (parseNode) => new GroupCollectionResponse());
                groups.AddRange(listgroups.Value);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Iterate through collection of groups                                                                   //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var group in groups)
            {
                var StringTeamsOptions = group.AdditionalData["resourceProvisioningOptions"].ToString();
                var CleanStringTeamsOptions = Regex.Replace(StringTeamsOptions, "[^a-zA-Z]", string.Empty);

                if (CleanStringTeamsOptions == "Team")
                {
                    if (exceptionGroupsArray.Contains(group.Id) == false)
                    {
                        var channels = await graphClient.Teams[group.Id].Channels.GetAsync();
                        var url = "";

                        foreach (var channel in channels.Value)
                        {
                            if (channel.DisplayName == "General")
                            {
                                url = "https://teams.microsoft.com/_#/l/team/" + channel.Id + "/conversations?groupId=" + group.Id + "&tenantId=" + tenantid;
                            }
                        };

                        // If no General channel found, take first channel
                        if (url == "")
                        {
                            url = "https://teams.microsoft.com/_#/conversations/" + channels.Value[0].DisplayName + "?threadId=" + channels.Value[0].Id;
                        }
                        
                        CreateList.Add(new CreateItem {Url = url, Id = group.Id});

                        foreach (var item in items)
                        {
                            //compare group id to the sharepoint list
                            if (item.Fields.AdditionalData["TeamsID"].ToString() == group.Id)
                            {
                                //compare the url
                                if (item.Fields.AdditionalData["Teamslink"].ToString() != url)
                                {
                                    //add to the list to be updated
                                    item.Fields.AdditionalData["Teamslink"] = url;
                                    UpdateList.Add(item);
                                }

                                //remove from the items collection
                                items.Remove(item);

                                var item1 = CreateList.SingleOrDefault(x => x.Id == group.Id);
                                CreateList.Remove(item1);
                                break;
                            }
                        }
                    }
                }
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Update items in UpdateList                                                                             //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var item in UpdateList)
            {
                log.LogInformation($"Updated item.Id: {item.Id}");
                var Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"TeamsID", item.Fields.AdditionalData["TeamsID"]},
                        {"Teamslink", item.Fields.AdditionalData["Teamslink"]}
                    }
                };

                await graphClient.Sites[siteId].Lists[listId].Items[item.Id].Fields.PatchAsync(Fields);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Add items in CreateList                                                                                //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var item in CreateList)
            {
                log.LogInformation($"Created item.Id: {item.Id}");
                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"TeamsID", item.Id},
                            {"Teamslink", item.Url}
                        }
                    }
                };

                await graphClient.Sites[siteId].Lists[listId].Items.PostAsync(listItem);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Delete remaining items                                                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var item in items)
            {
                log.LogInformation($"Deleted item.Id: {item.Id}");
                await graphClient.Sites[siteId].Lists[listId].Items[item.Id].DeleteAsync();
            }
        }
    }
}