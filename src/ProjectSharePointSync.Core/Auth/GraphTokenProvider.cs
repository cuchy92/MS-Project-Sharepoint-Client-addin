using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using ProjectSharePointSync.Core.Diagnostics;

namespace ProjectSharePointSync.Core.Auth
{
    /// <summary>
    /// Acquires delegated Microsoft Graph tokens for the signed-in user using MSAL.
    /// Tokens are cached on disk (encrypted with DPAPI), so the user signs in only once.
    /// </summary>
    public sealed class GraphTokenProvider
    {
        /// <summary>
        /// Sites.Manage.All is needed to create lists and columns.
        /// If the lists are provisioned up front, Sites.ReadWrite.All is enough.
        /// </summary>
        public static readonly string[] DefaultScopes = { "https://graph.microsoft.com/Sites.Manage.All" };

        private readonly IPublicClientApplication _app;
        private readonly string[] _scopes;
        private readonly SemaphoreSlim _interactiveGate = new SemaphoreSlim(1, 1);
        private readonly Lazy<Task> _cacheInit;

        public GraphTokenProvider(string tenantId, string clientId, string[] scopes = null)
        {
            if (string.IsNullOrWhiteSpace(clientId)) throw new ArgumentException("Client ID is required.", "clientId");
            if (string.IsNullOrWhiteSpace(tenantId)) tenantId = "organizations";

            TenantId = tenantId.Trim();
            ClientId = clientId.Trim();
            _scopes = scopes ?? DefaultScopes;

            _app = PublicClientApplicationBuilder
                .Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                .WithRedirectUri("http://localhost")
                .WithClientName("ProjectSharePointSync")
                .Build();

            _cacheInit = new Lazy<Task>(InitializeCacheAsync, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        public string TenantId { get; private set; }

        public string ClientId { get; private set; }

        /// <summary>
        /// Window handle used as the parent of the sign-in UI (optional).
        /// </summary>
        public IntPtr ParentWindow { get; set; }

        /// <summary>
        /// Returns a valid access token. Uses the cache first; falls back to an interactive
        /// browser sign-in when <paramref name="allowInteractive"/> is true.
        /// </summary>
        public async Task<string> GetAccessTokenAsync(bool allowInteractive, CancellationToken cancellationToken)
        {
            await _cacheInit.Value.ConfigureAwait(false);

            var account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            if (account != null)
            {
                try
                {
                    var silent = await _app.AcquireTokenSilent(_scopes, account)
                        .ExecuteAsync(cancellationToken).ConfigureAwait(false);
                    return silent.AccessToken;
                }
                catch (MsalUiRequiredException ex)
                {
                    Log.Info("Silent token acquisition failed, interactive sign-in required: " + ex.ErrorCode);
                }
            }

            if (!allowInteractive)
            {
                throw new AuthenticationRequiredException("Sign-in to SharePoint is required. Use 'Sync now' in the SharePoint ribbon tab.");
            }

            await _interactiveGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                // Another caller may have signed in while we were waiting.
                account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
                if (account != null)
                {
                    try
                    {
                        var again = await _app.AcquireTokenSilent(_scopes, account)
                            .ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        return again.AccessToken;
                    }
                    catch (MsalUiRequiredException)
                    {
                    }
                }

                var builder = _app.AcquireTokenInteractive(_scopes)
                    .WithPrompt(Prompt.SelectAccount)
                    .WithUseEmbeddedWebView(false);

                if (account != null)
                {
                    builder = builder.WithAccount(account);
                }

                if (ParentWindow != IntPtr.Zero)
                {
                    builder = builder.WithParentActivityOrWindow(ParentWindow);
                }

                var result = await builder.ExecuteAsync(cancellationToken).ConfigureAwait(false);
                Log.Info("Signed in as " + result.Account.Username);
                return result.AccessToken;
            }
            finally
            {
                _interactiveGate.Release();
            }
        }

        public async Task SignOutAsync()
        {
            await _cacheInit.Value.ConfigureAwait(false);
            foreach (var account in await _app.GetAccountsAsync().ConfigureAwait(false))
            {
                await _app.RemoveAsync(account).ConfigureAwait(false);
            }
        }

        public async Task<string> GetSignedInUserAsync()
        {
            await _cacheInit.Value.ConfigureAwait(false);
            var account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            return account == null ? null : account.Username;
        }

        private async Task InitializeCacheAsync()
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ProjectSharePointSync");
                Directory.CreateDirectory(dir);

                var props = new StorageCreationPropertiesBuilder("msal_" + ClientId + ".cache", dir).Build();
                var helper = await MsalCacheHelper.CreateAsync(props).ConfigureAwait(false);
                helper.RegisterCache(_app.UserTokenCache);
            }
            catch (Exception ex)
            {
                // Without a persistent cache the add-in still works, the user just signs in once per session.
                Log.Error("Could not initialize the persistent token cache.", ex);
            }
        }
    }

    public sealed class AuthenticationRequiredException : Exception
    {
        public AuthenticationRequiredException(string message) : base(message)
        {
        }
    }
}
