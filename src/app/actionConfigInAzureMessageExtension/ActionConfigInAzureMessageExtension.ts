import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { TaskModuleRequest, TaskModuleContinueResponse } from "botbuilder";
import { AppConfigurationClient } from "@azure/app-configuration";
// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/actionConfigInAzureMessageExtension/config.html")
@PreventIframe("/actionConfigInAzureMessageExtension/action.html")
export default class ActionConfigInAzureMessageExtension implements IMessagingExtensionMiddlewareProcessor {
    public async onFetchTask(context: TurnContext, value: MessagingExtensionQuery): Promise<MessagingExtensionResult | TaskModuleContinueResponse> {
        if (false) { // !value.state TODO: implement logic when config is persisted
            return Promise.resolve<MessagingExtensionResult>({
                type: "config", // use "config" or "auth" here
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: `https://${process.env.HOSTNAME}/actionConfigInAzureMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`,
                            title: "Configuration"
                        }
                    ]
                }
            });
        }

        return Promise.resolve<TaskModuleContinueResponse>({
            type: "continue",
            value: {
                title: "Input form",
                url: `https://${process.env.HOSTNAME}/actionConfigInAzureMessageExtension/action.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`,
                height: "medium"
            }
        });
    }

    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: TaskModuleRequest): Promise<MessagingExtensionResult> {
        const card = CardFactory.adaptiveCard(
            {
              type: "AdaptiveCard",
              body: [
                {
                  type: "ColumnSet",
                  columns: [
                      {
                        type: "Column",
                        width: 25,
                        items: [
                          {
                            type: "Image",
                            url: `https://${process.env.HOSTNAME}/assets/icon.png`,
                            style: "Person"
                          }
                        ]
                      },
                      {
                        type: "Column",
                        width: 75,
                        items: [
                          {
                            type: "TextBlock",
                            text: value.data.doc.name,
                            size: "Large",
                            weight: "Bolder"
                          },
                          {
                            type: "TextBlock",
                            text: `Author: ${value.data.doc.author}`
                          },
                          {
                            type: "TextBlock",
                            text: `Modified: ${value.data.doc.modified}`
                          }
                        ]
                      }
                  ]
                }
              ],
              actions: [
                {
                    type: "Action.OpenUrl",
                    title: "View",
                    url: value.data.doc.url
                }
              ],
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              version: "1.0"
            });
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [card]
            } as MessagingExtensionResult);
    }



    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
        const client = new AppConfigurationClient(connectionString);
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
              siteID = process.env.SITE_ID!;
          }
          if (listID === "") {
              listID = process.env.LIST_ID!;
          }
        }
        return Promise.resolve({
            title: "Action Config in Azure Configuration",
            value: `https://${process.env.HOSTNAME}/actionConfigInAzureMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}&siteID=${siteID}&listID=${listID}`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = JSON.parse(context.activity.value.state);
        log(`New setting: ${setting}`);
        const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
        const client = new AppConfigurationClient(connectionString);
        const siteID = setting.siteID;
        const listID = setting.listID;
        if (siteID) {
          await client.setConfigurationSetting({ key: "SiteID", value: siteID });
        }
        if (listID) {
          await client.setConfigurationSetting({ key: "ListID", value: listID });
        }
        return Promise.resolve();
    }
}
