using Microsoft.Azure.KeyVault;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AzureResourceReport.Models
{
    public class KeyVaultCache
    {
        public KeyVaultCache(string baseUri, string clientId, string clientSecret)
        {
            BaseUri = baseUri;
            ClientId = clientId;
            ClientSecret = clientSecret;
        }

        public static string BaseUri { get; set; }

        public static string ClientId { get; set; }

        public static string ClientSecret { get; set; }

        private static KeyVaultClient _KeyVaultClient = null;

        private static Dictionary<string, string> SecretsCache = new Dictionary<string, string>();

        public async Task<string> GetCachedSecret(string secretName)
        {
            if (!SecretsCache.ContainsKey(secretName))
            {
                if (_KeyVaultClient is null)
                {
                    _KeyVaultClient = new KeyVaultClient(async (authority, resource, scope) =>
                    {
                        var adCredential = new ClientCredential(ClientId, ClientSecret);
                        var authenticationContext = new AuthenticationContext(authority, null);
                        return (await authenticationContext.AcquireTokenAsync(resource, adCredential)).AccessToken;
                    });
                }

                var secretBundle = await _KeyVaultClient.GetSecretAsync($"{BaseUri}{secretName}").ConfigureAwait(false);
                SecretsCache.Add(secretName, secretBundle.Value);
            }

            return SecretsCache.ContainsKey(secretName) ? SecretsCache[secretName] : string.Empty;
        }
    }
}