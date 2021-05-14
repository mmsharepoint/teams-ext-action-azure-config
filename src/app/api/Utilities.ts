import { AppConfigurationClient } from "@azure/app-configuration";
import { DefaultAzureCredential, EnvironmentCredential, ManagedIdentityCredential } from "@azure/identity";
import { Config } from "../../model/Config";

export default class Utilities {
  public static async retrieveConfig(): Promise<Config> {
    const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
    // const credential = new DefaultAzureCredential();
    // console.log(credential);
    // const client = new AppConfigurationClient(connectionString, credential);
    let client: AppConfigurationClient;
    if (process.env.AZURE_CLIENT_SECRET) {
      const credential = new EnvironmentCredential();
      client = new AppConfigurationClient(connectionString, credential);
    }
    else {
      const credential = new ManagedIdentityCredential();
      console.log("Managed Identity used");
      console.log(credential);
      client = new AppConfigurationClient(connectionString, credential);
    }
    let siteID = "";
    let listID = "";
    try {
      const siteSetting = await client.getConfigurationSetting({ key: "SiteID"});
      siteID = siteSetting.value!;
      const listSetting = await client.getConfigurationSetting({ key: "ListID"});
      listID = listSetting.value!;
    }
    catch(error) {
      if (siteID === "") {
        //  siteID = process.env.SITE_ID!;
      }
      if (listID === "") {
        //  listID = process.env.LIST_ID!;
      }
    }
    return Promise.resolve({ SiteID: siteID, ListID: listID})
  }

  public static async saveConfig(newConfig: Config) {
    const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
    const client = new AppConfigurationClient(connectionString);
    if (newConfig.SiteID) {
      await client.setConfigurationSetting({ key: "SiteID", value: newConfig.SiteID });
    }
    if (newConfig.ListID) {
      await client.setConfigurationSetting({ key: "ListID", value: newConfig.ListID });
    }
  }
}