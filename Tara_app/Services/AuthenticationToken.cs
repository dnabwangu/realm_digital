namespace Tara_app.Services
{
    class AuthenticationToken
    {
        public string access_token { get; set; }
        public double expires_in { get; set; }
        public string token_type { get; set; }
    }
}
