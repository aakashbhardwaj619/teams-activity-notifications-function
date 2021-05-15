using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;

namespace TeamsAzDemo
{
    public class MSALAuthenticationProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _clientApplication;
        private string[] _scopes;

        public MSALAuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
        {
            _clientApplication = clientApplication;
            _scopes = scopes;
        }

        /// <summary>
        /// Update HttpRequestMessage with credentials
        /// </summary>
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var token = await GetTokenAsync();
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
        }

        /// <summary>
        /// Acquire Token 
        /// </summary>
        public async Task<string> GetTokenAsync()
        {
            AuthenticationResult authResult = null;
            authResult = await _clientApplication.AcquireTokenForClient(_scopes)
                                .ExecuteAsync();
            return authResult.AccessToken;
        }
    }
}