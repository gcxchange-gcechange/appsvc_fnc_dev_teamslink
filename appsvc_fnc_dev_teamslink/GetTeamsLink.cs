using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace appsvc_fnc_dev_teamslink
{
    public class GetTeamsLink
    {
        private readonly ILogger<GetTeamsLink> _logger;
        public GetTeamsLink(ILogger<GetTeamsLink> logger)
        {
            _logger = logger;
        }

        [Function("GetTeamsLink")]
        //Timezone UTC universal
        public async Task Run([TimerTrigger("0 0 10-21/2 * * 1-5")] TimerInfo myTimer)
        {
            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.UtcNow}");

            if (myTimer.ScheduleStatus is not null)
            {
                _logger.LogInformation("Next timer schedule at: {nextSchedule}", myTimer.ScheduleStatus.Next);
            }

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            var exceptionGroupsArray = config["exceptionGroupsArray"];
            var siteId = config["siteId"];
            var listId = config["listId"];
            var tenantid = config["tenantid"];

            Auth auth = new Auth();

            var graphClient = auth.graphAuth(_logger);
            var UpdateList = new List<ListItem>();

            List<CreateItem> CreateList = new List<CreateItem>();

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Get items from TeamsLink list                                                                          //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            var listitems = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=TeamsID,Teamslink)" };
                requestConfiguration.QueryParameters.Top = 999;
            });

            List<ListItem> items = new List<ListItem>();
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

            _logger.LogInformation($"Total items in TeamsLink list: {items.Count}");

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Get groups from tenant                                                                                 //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///
            var listgroups = await graphClient.Groups.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id,resourceProvisioningOptions" };

                //GET /groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')
                requestConfiguration.QueryParameters.Filter = "resourceProvisioningOptions/Any(x:x eq 'Team')";

                requestConfiguration.QueryParameters.Top = 999;
            });

            var groups = new List<Microsoft.Graph.Models.Group>();
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

            _logger.LogInformation($"Total groups in tenant: {groups.Count}");

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Iterate through collection of groups                                                                   //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var group in groups)
            {
                if (exceptionGroupsArray.Contains(group.Id) == false)
                {
                    var channels = await graphClient.Teams[group.Id].Channels.GetAsync();
                    var url = "";

                    foreach (var channel in channels.Value)
                    {
                        if (channel.DisplayName == "General")
                        {
                            url = "https://teams.microsoft.com/#/l/team/" + channel.Id + "/conversations?groupId=" + group.Id + "&tenantId=" + tenantid;
                        }
                    };

                    // if no General channel found, take first channel
                    if (url == "")
                    {
                        url = "https://teams.microsoft.com/#/l/conversations/" + channels.Value[0].DisplayName + "?threadId=" + channels.Value[0].Id;
                    }

                    CreateList.Add(new CreateItem { Url = url, ID = group.Id });

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

                            var item1 = CreateList.SingleOrDefault(x => x.ID == group.Id);
                            CreateList.Remove(item1);
                            break;
                        }
                    }
                }
                else
                {
                    _logger.LogInformation($"Skipping Group ID {group.Id} as it is in the exception list.");
                }
            }

            _logger.LogInformation($"Total items to be updated: {UpdateList.Count}");
            _logger.LogInformation($"Total items to be created: {CreateList.Count}");
            _logger.LogInformation($"Total items to be deleted: {items.Count}");

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Update items in UpdateList                                                                             //
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (var item in UpdateList)
            {
                _logger.LogInformation($"Updated item.Id: {item.Id}");
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
                _logger.LogInformation($"Created item.Id: {item.ID}");
                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"TeamsID", item.ID},
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
                _logger.LogInformation($"Deleted item.Id: {item.Id}");
                await graphClient.Sites[siteId].Lists[listId].Items[item.Id].DeleteAsync();
            }
        }
    }
}