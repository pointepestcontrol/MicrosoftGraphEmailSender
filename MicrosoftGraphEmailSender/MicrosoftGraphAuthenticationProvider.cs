using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphEmailSender
{
    public class MicrosoftGraphAuthenticationProvider : IAuthenticationProvider
    {

        private readonly string clientId;
        private readonly string clientSecret;
        private readonly string[] appScopes;
        private readonly string tenantId;

        public MicrosoftGraphAuthenticationProvider(string clientId, string clientSecret, string[] appScopes, string tenantId)
        {
            this.clientId = clientId;
            this.clientSecret = clientSecret;
            this.appScopes = appScopes;
            this.tenantId = tenantId;
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var clientApplication = ConfidentialClientApplicationBuilder.Create(this.clientId)
                .WithClientSecret(this.clientSecret)
                .WithTenantId(this.tenantId)
                .Build();

            var result = await clientApplication.AcquireTokenForClient(this.appScopes).ExecuteAsync();

            request.Headers.Add("Authorization", result.CreateAuthorizationHeader());
        }
    }
}