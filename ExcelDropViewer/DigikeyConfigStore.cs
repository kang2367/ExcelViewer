namespace ExcelDropViewer
{
    internal sealed class DigikeyConfig
    {
        public string ClientId { get; set; } = string.Empty;
        public string ClientSecret { get; set; } = string.Empty;
    }

    internal static class DigikeyConfigStore
    {
        private const string SectionName = "DIGIKEY";
        private const string ClientIdKey = "CLIENT_ID";
        private const string ClientSecretKey = "CLIENT_SECRET";

        public static DigikeyConfig Load()
        {
            return new DigikeyConfig
            {
                ClientId = ConfigIniFile.ReadValue(SectionName, ClientIdKey) ?? string.Empty,
                ClientSecret = ConfigIniFile.ReadValue(SectionName, ClientSecretKey) ?? string.Empty
            };
        }

        public static void Save(DigikeyConfig config)
        {
            ConfigIniFile.WriteSection(SectionName, new Dictionary<string, string>
            {
                [ClientIdKey] = config.ClientId,
                [ClientSecretKey] = config.ClientSecret
            });
        }
    }
}
