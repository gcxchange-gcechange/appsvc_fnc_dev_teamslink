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

namespace appsvc_fnc_dev_teamslink
{
    public static class GetTeamsLink
    {
        [FunctionName("GetTeamsLink")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder()

           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");
            var exceptionGroupsArray = config["exceptionGroupsArray"];
            var siteId = config["siteId"];
            var listId = config["listId"];

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            var UpdateList = new ListItemsCollectionPage();
            var CreateList = new ListItemsCollectionPage();


            var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("expand", "fields(select=Title,TeamsID,Teamslink)")
                };

            var AllTeamsItems = await graphClient.Sites[siteId].Lists[listId].Items
                .Request(queryOptions)
                .GetAsync();

            var groups = await graphClient.Groups
                .Request()
                .Select("id,resourceProvisioningOptions")
                .GetAsync();

            foreach (var group in groups)
            {
                if (group.AdditionalData["resourceProvisioningOptions"] == "Teams")
                {
                    string groupid = group.Id;
                    string group_name = group.DisplayName;
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
                                url = "https://teams.microsoft.com/_#/conversations/General?threadId=" + channel.Id;
                            }
                        };

                        // If no General channel found, take first channel
                        if (url == "")
                        {
                            url = "https://teams.microsoft.com/_#/conversations/" + channels[0].DisplayName + "?threadId=" + channels[0].Id;
                        }

                        //check if part of list
                        foreach (var item in AllTeamsItems)
                        {
                            //compare group id to the sharepoint list
                            if (item.AdditionalData["TeamsID"] == group.Id)
                            {
                                //compare the url
                                if (item.AdditionalData["Teamslink"] != url)
                                {
                                    //add to the list to be update
                                    UpdateList.Add(item);
                                }
                                //remove from the all list
                                AllTeamsItems.Remove(item);
                            }
                            else
                            {
                                //Group id is not part of the list
                                //Need to create a new item to the list and remove from the allItems
                                CreateList.Add(item);
                                AllTeamsItems.Remove(item);
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
                        {"Title", item.AdditionalData["Title"]},
                        {"TeamsID", item.AdditionalData["TeamsID"]},
                        {"Teamslink", item.AdditionalData["Teamslink"]}
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
                            {"Title", item.AdditionalData["Title"]},
                            {"TeamsID", item.AdditionalData["TeamsID"]},
                            {"Teamslink", item.AdditionalData["Teamslink"]}
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
                await graphClient.Sites[siteId].Lists[listId].Items[item.Id].Fields
                    .Request()
                    .DeleteAsync();
            }


            string responseMessage = "Success";
            return new OkObjectResult(responseMessage);
        }
    }
}
