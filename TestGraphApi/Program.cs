using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace TestGraphApi
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var config = new ConfigurationBuilder()
             .SetBasePath(System.IO.Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json").Build();

            var graphConfig = new GraphConfig();
            config.GetSection("GraphConfigs").Bind(graphConfig);


            IConfidentialClientApplication app;
            app = ConfidentialClientApplicationBuilder.Create(graphConfig.ClientId)
                                                      .WithTenantId(graphConfig.TenantId)
                                                      .WithClientSecret(graphConfig.ClientSecret)
                                                      .Build();

            var authProvider = new ClientCredentialProvider(app);
            var graphServiceClient = new GraphServiceClient(authProvider);

            // know issue getting places directly: https://github.com/microsoftgraph/msgraph-beta-sdk-dotnet/issues/97
            // var room = await graphServiceClient.Places.Request().GetAsync();

            var roomsRequestUrl = graphServiceClient.Places.AppendSegmentToRequestUrl("microsoft.graph.room");
            var placesRequest = new GraphServicePlacesCollectionRequest(roomsRequestUrl, graphServiceClient, null);
            var places = await placesRequest.GetAsync();

            var scheduleCollection = await GetRoomAvailability(graphServiceClient, places);
            foreach (var place in places)
            {
                var schedule = scheduleCollection.FirstOrDefault(x => x.ScheduleId == place.AdditionalData["emailAddress"].ToString());
                var availability = string.Empty;
                if (schedule != null)
                {
                    availability = schedule.AvailabilityView == "0" ? FreeBusyStatus.Free.ToString() : FreeBusyStatus.Busy.ToString();
                }
                Console.WriteLine($"Room - {place.DisplayName}  Availability   - {availability}    FloorId - {place.AdditionalData["floorNumber"]}     BuildingId - {place.AdditionalData["building"]}");
            }
        }

        private async static Task<ICalendarGetScheduleCollectionPage> GetRoomAvailability(GraphServiceClient graphServiceClient, IGraphServicePlacesCollectionPage places)
        {
            var schedules = places.Select(x => x.AdditionalData["emailAddress"].ToString()).ToList();

            var startTime = new DateTimeTimeZone
            {
                DateTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "AUS Eastern Standard Time"
            };

            var endTime = new DateTimeTimeZone
            {
                DateTime = DateTime.Now.AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "AUS Eastern Standard Time"
            };

            var availabilityViewInterval = 60;

            var meetingRoomSchedules = await graphServiceClient.Users[places.First().AdditionalData["emailAddress"].ToString()].Calendar
                .GetSchedule(schedules, endTime, startTime, availabilityViewInterval)
                .Request()
                .Header("Prefer", "outlook.timezone=\"AUS Eastern Standard Time\"")
                .PostAsync();

            return meetingRoomSchedules;
        }        
    }
}


