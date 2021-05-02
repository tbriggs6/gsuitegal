using Microsoft.Win32;
using System;

namespace GSuiteGAL
{

    public class Config
    {
        private const string registryRoot = "Software\\GoogleGAL";
        private const string installPathKey = "install_path";
        private const string credentialKey = "cred_file";
        private const string tokenKey = "token_file";
        private const string lastSyncKey = "lastsync";
        private const string syncPeriodKey = "sync_period";

        private RegistryKey configKey = null;

        //public int Age { get; set; }
        public string installPathName { get; } = null;
        public String credentialFileName { get; } = null;
        public String tokenFileName { get; } = null;
        public int lastSync { get; set; } = 0;
        public int syncPeriod { get; } = 0;


        public Config()
        {
            configKey = OpenOrCreateSubKey();
            credentialFileName = (string)configKey.GetValue(credentialKey);
            tokenFileName = (string)configKey.GetValue(tokenKey);
            lastSync = (int)configKey.GetValue(lastSyncKey);
            syncPeriod = (int)configKey.GetValue(syncPeriodKey);
            installPathName = (string)configKey.GetValue(installPathKey);
        }

        public void SetLastSync(int time)
        {
            lastSync = time;
            RegistryKey configKey = Registry.CurrentUser.OpenSubKey(registryRoot, true);
            configKey.SetValue(lastSyncKey, time);
        }

        private static RegistryKey OpenOrCreateSubKey()
        {
            try
            {
                RegistryKey configKey = Registry.CurrentUser.OpenSubKey(registryRoot);

                // create subkey, it doesn't exist.
                if (configKey == null)
                {
                    configKey = Registry.CurrentUser.CreateSubKey(registryRoot);
                    if (configKey == null)
                        throw new SystemException("Registry failed.");

                    configKey.SetValue(installPathKey, "c:\\users\\tbriggs\\source\\repos\\GSuiteGAL\\GSuiteGAL");
                    configKey.SetValue(credentialKey, "credentials.json");
                    configKey.SetValue(tokenKey, "tokens.json");
                    configKey.SetValue(lastSyncKey, 0);
                    configKey.SetValue(syncPeriodKey, 86400); // once a day (in seconds)
                }
                return configKey;
            }
            catch (SystemException e)
            {
                //        EventLog.WriteEntry("GoogleGAL", "System exception " + e.ToString(),
                //             EventLogEntryType.Error, 1000, 1000, null);
                throw e;  // re-throw it
            }

        }


    }
}
