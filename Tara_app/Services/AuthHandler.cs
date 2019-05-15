using Newtonsoft.Json;
using System;
using System.Net;

namespace Tara_app.Services
{
    public sealed class AuthHandler
    {
        public const string authUrl = "https://api.ai-spark.com/v1/authenticate";
        private AuthForm authForm = null;

        private AuthHandler()
        {
        }

        private static readonly Lazy<AuthHandler> lazy = new Lazy<AuthHandler>(() => new AuthHandler());
        public static AuthHandler Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        public bool IsUserAuthenticated()
        {
            DateTime expirationDate = GetTokeExpiration();
            DateTime nowDate = DateTime.Now;

            var dateComparator = DateTime.Compare(expirationDate, nowDate);

            if (dateComparator <= 0)
            {
                return false;
            } else
            {
                return true;
            }
        }

        public DateTime GetTokeExpiration()
        {
            var expiration = Properties.Settings.Default.TokenExpiration;
            try
            {
                DateTime expirationDate = Convert.ToDateTime(expiration);

                return expirationDate;
            }
            catch
            {
                return DateTime.Now;
            }
        }

        public void SetTokenExpiration(double expiration)
        {
            TimeSpan expirationSpan = TimeSpan.FromSeconds(expiration);
            DateTime dateNow = DateTime.Now;

            DateTime expirationDate = dateNow + expirationSpan;

            Properties.Settings.Default.TokenExpiration = expirationDate.ToString();
            Properties.Settings.Default.Save();
        }

        public string GetAuthHeader()
        {
            var header = Properties.Settings.Default.AuthToken;
            return header;
        }

        private void StoreHeader(string header)
        {
            Properties.Settings.Default.AuthToken = header;
            Properties.Settings.Default.Save();
        }

        public string GetAuthUser()
        {
            var user = Properties.Settings.Default.AuthUser;
            return user;
        }

        private void StoreAuthUser(string user)
        {
            Properties.Settings.Default.AuthUser = user;
            Properties.Settings.Default.Save();
        }

        public void ShowAuthForm()
        {
            if (this.authForm == null)
            {
                this.authForm = new AuthForm();
            }

            authForm.Show();
        }

        public void CloseAuthForm()
        {
            if (this.authForm != null)
            {
               authForm.Close();
            }
        }

        public void ShowAuthError(string errorText)
        {
            if (this.authForm != null)
            {
                authForm.ShowError(errorText);
            }
        }

        public void StartAuthenticationProcess()
        {

        }

        private void AuthenticationSuccess(string userName)
        {
            StoreAuthUser(userName);

            Globals.Ribbons.Ribbon1.ShowAuthenticatedLayout();

            CloseAuthForm();
        }

        public void Logout()
        {
            StoreHeader("");
            StoreAuthUser("");
            SetTokenExpiration(0);

            Globals.Ribbons.Ribbon1.ShowUnAuthenticatedLayout();
        }

        public void authenticateUser(string userName, string password)
        {
            using (WebClient client = new WebClient())
            {
                var userObject = new { email = userName, password = password };
                var serializedUserObject = JsonConvert.SerializeObject(userObject);

                client.Headers.Add(HttpRequestHeader.ContentType, "application/json");
                var responsebody = client.UploadString(authUrl, "POST", serializedUserObject);

                if (!responsebody.ToLower().Contains("error"))
                {
                    DeserializeToken(responsebody);
                    AuthenticationSuccess(userName);
                } else
                {
                    var errorMessage = DeserializeError(responsebody);

                    ShowAuthError(errorMessage);
                }

            }

        }

        private void DeserializeToken(string token)
        {
            var deserializedObject = JsonConvert.DeserializeObject<AuthenticationToken>(token);
            var accessToken = deserializedObject.access_token;
            var tokenType = deserializedObject.token_type;

            var authHeader = tokenType + " " + accessToken;
            var expiration = deserializedObject.expires_in;

            StoreHeader(authHeader);
            SetTokenExpiration(expiration);
        }

        private string DeserializeError(string token)
        {
            var deserializedObject = JsonConvert.DeserializeObject < AuthenticationError>(token);
            var errorMessage = deserializedObject.error_description;

            return errorMessage;
        }
    }
}
