using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace appsvc_fnc_dev_teamslink
{
    public static class GetTeamsLink
    {
        // Timezone UTC universal

        // previous: "0 0 10-21/2 * * 1-5"
        // At minute 0 past every 2nd hour from 10 through 21 on every day-of-week from Monday through Friday.

        // current: "0 0 6 * * 1"
        // At 06:00 on Monday.

        [FunctionName("GetTeamsLink")]
        public static async Task Run([TimerTrigger("0 0 6 * * 1")] TimerInfo myTimer, ILogger log)
        {
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
            var UpdateList = new ListItemsCollectionPage();
            List<CreateItem> CreateList = new List<CreateItem>();

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("expand", "fields(select=TeamsID,Teamslink)")
            };
            List<ListItem> items = new List<ListItem>();
            var AllTeamsItems = await graphClient.Sites[siteId].Lists[listId].Items
            .Request(queryOptions)
            .Top(999)
            .GetAsync();

            items.AddRange(AllTeamsItems.CurrentPage.OfType<ListItem>());
            // fetch next page
            while (AllTeamsItems.NextPageRequest != null)
            {
                AllTeamsItems = await AllTeamsItems.NextPageRequest.GetAsync();
                items.AddRange(AllTeamsItems.CurrentPage.OfType<ListItem>());
            }



            var groups = new List<Microsoft.Graph.Group>();
            var listgroups = await graphClient.Groups
                .Request()
                .Select("id,resourceProvisioningOptions")
                .Top(999)
                .GetAsync();

            groups.AddRange(listgroups.CurrentPage.OfType<Microsoft.Graph.Group>());
            // fetch next page
            while (listgroups.NextPageRequest != null)
            {
                listgroups = await listgroups.NextPageRequest.GetAsync();
                groups.AddRange(listgroups.CurrentPage.OfType<Microsoft.Graph.Group>());
            }

            foreach (var group in groups)
            {
                var StringTeamsOptions = group.AdditionalData["resourceProvisioningOptions"].ToString();
                var CleanStringTeamsOptions = Regex.Replace(StringTeamsOptions, "[^a-zA-Z]", string.Empty);

                if (CleanStringTeamsOptions == "Team")
                {
                    if (exceptionGroupsArray.Contains(group.Id) == false)
                    {
                        var channels = await graphClient.Teams[group.Id].Channels
                        .Request()
                        .GetAsync();

                        var url = "";

                        foreach (var channel in channels)
                        {
                            if (channel.DisplayName == "General")
                            {
                                url = "https://teams.microsoft.com/_#/l/team/" + channel.Id + "/conversations?groupId=" + group.Id + "&tenantId=" + tenantid;
                            }
                        };

                        // If no General channel found, take first channel
                        if (url == "")
                        {
                            url = "https://teams.microsoft.com/_#/conversations/" + channels[0].DisplayName + "?threadId=" + channels[0].Id;
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
                                    //add to the list to be update
                                    item.Fields.AdditionalData["Teamslink"] = url;
                                    UpdateList.Add(item);
                                }
                                //remove from the all list

                                AllTeamsItems.Remove(item);
                                var item1 = CreateList.SingleOrDefault(x => x.ID == group.Id);
                                CreateList.Remove(item1);
                                break;
                            }
                        }
                    }
                }
            }

            //function to update all item
            foreach (var item in UpdateList)
            {
                var Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"TeamsID", item.Fields.AdditionalData["TeamsID"]},
                        {"Teamslink", item.Fields.AdditionalData["Teamslink"]}
                    }
                };

                await graphClient.Sites[siteId].Lists[listId].Items[item.Id].Fields
                    .Request()
                    .UpdateAsync(Fields);
            }

            //Function to create all item
            foreach (var item in CreateList)
            {
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
                await graphClient.Sites[siteId].Lists[listId].Items
                    .Request()
                    .AddAsync(listItem);
            }

            //Function to delete all item from all list
            foreach (var item in AllTeamsItems)
            {
                await graphClient.Sites[siteId].Lists[listId].Items[item.Id]
                .Request()
                .DeleteAsync();
            }

            string responseMessage = "Success";
            //return new OkObjectResult(responseMessage);
        }
    }
}
