/*
Microsoft provides programming examples for illustration only, without 
warranty either expressed or implied, including, but not limited to, the
implied warranties of merchantability and/or fitness for a particular 
purpose.  
This sample assumes that you are familiar with the programming language
being demonstrated and the tools used to create and debug procedures. 
Microsoft support professionals can help explain the functionality of a
particular procedure, but they will not modify these examples to provide
added functionality or construct procedures to meet your specific needs. 
If you have limited programming experience, you may want to contact a 
Microsoft Certified Partner or the Microsoft fee-based consulting line 
at (800) 936-5200. 
*/

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Documents;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.IdentityModel.Protocols;

namespace CosmosAAD
{
    public static class CosmosTrigger
    {
        [FunctionName("CosmosTrigger")]
        public static async Task RunAsync([CosmosDBTrigger(
            databaseName: "db",
            collectionName: "in",
            ConnectionStringSetting = "CosmosConnectionString",
            LeaseCollectionName = "leases")]IReadOnlyList<Document> input,
            [CosmosDB(
                databaseName: "db",
                collectionName: "out",
                ConnectionStringSetting = "CosmosConnectionString")] IAsyncCollector<ProcessedDoc> output,
            ILogger log)
        {
            if (input != null && input.Count > 0)
            {
                //Get the from email
                string strFromPhone                                     = input[0].GetPropertyValue<string>("from");
                log.LogInformation("The email from: " + strFromPhone);

                string strEMail                                         = await GetUserInfoByPhoneAsync(strFromPhone);

                //Create new doc
                ProcessedDoc md                                         = new ProcessedDoc();
                md.id                                                   = input[0].Id;
                md.EMailFrom                                            = strEMail;
                await output.AddAsync(md);
            }
        }


        private static async Task<string> GetUserInfoByPhoneAsync(string strPhone)
        {
            var scopes                                                  = new[] { "https://graph.microsoft.com/.default" };

            // Values from app registration
            var clientId                                                = Environment.GetEnvironmentVariable("ClientID", EnvironmentVariableTarget.Process);
            var tenantId                                                = Environment.GetEnvironmentVariable("TenantId", EnvironmentVariableTarget.Process);
            var clientSecret                                            = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);

            var options                                                 = new ClientSecretCredentialOptions
            {
                AuthorityHost                                           = AzureAuthorityHosts.AzurePublicCloud,
            };

            var clientSecretCredential                                  = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

            var graphClient                                             = new GraphServiceClient(clientSecretCredential, scopes);

            var result                                                  = await graphClient.Users.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Count              = true;
                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");

                //Does not work
                //requestConfiguration.QueryParameters.Filter = "MobilePhone eq '"+strPhone+"'";
                //GH Issue
                //https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/2011
            });

            foreach (var item in result.Value)
            {
                if (item.MobilePhone == strPhone)
                    return item.UserPrincipalName; 
            }

            return string.Empty;
        }
    }

    /// <summary>
    /// This will be the output document format, TDB
    /// </summary>
    public class ProcessedDoc
    {
        public string id { get; set; }
        public string PartitionKey { get; set; }
        public string EMailFrom { get; set; }
    }
}
