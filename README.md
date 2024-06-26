---
page_type: sample
description: This sample demonstrates use of various meeting events and meeting participant events which are available in bot framework v4
products:
- office-teams
- office
- office-365
languages:
- csharp
extensions:
 contentType: samples
 createdDate: "11/10/2021 23:35:25 PM"
urlFragment: officedev-microsoft-teams-samples-meetings-events-csharp
---

# Realtime meeting events

Using this C# sample, a bot can receive real-time meeting events and meeting participant events.
For reference please check [Real-time Teams meeting events](https://docs.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/api-references?tabs=dotnet)
and [Real-time Teams meeting participant events](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis?branch=pr-8455&tabs=dotnet%2Cparticipant-join-event#get-participant-events)

The feature shown in this sample is currently available in public developer preview only.

## Included Features
* Bots
* Adaptive Cards
* RSC Permissions

## Interaction with app

![Meetings Events](MeetingEvents/Images/MeetingsEvents.gif)

## Try it yourself - experience the App in your Microsoft Teams client
Please find below demo manifest which is deployed on Microsoft Azure and you can try it yourself by uploading the app manifest (.zip file link below) to your teams and/or as a personal app. (Sideloading must be enabled for your tenant, [see steps here](https://docs.microsoft.com/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading)).

**Realtime meeting and participant events:** [Manifest](/samples/meetings-events/csharp/demo-manifest/Meetings-Events.zip)

## Prerequisites

- [.NET Core SDK](https://dotnet.microsoft.com/download) version 6.0

  ```bash
  # determine dotnet version
  dotnet --version
  ```
- Publicly addressable https url or tunnel such as [dev tunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started?tabs=windows) or [serveo.net](https://serveo.net/) latest version or [Tunnel Relay](https://github.com/OfficeDev/microsoft-teams-tunnelrelay) 

## Setup
> NOTE: if you want to have a fixed domain with Serveo, you should register first e.g.: `ssh -R mydomain:80:localhost:3978 serveo.net`, this will avoid you have a new URL everytime you run the command.

1) Setup for Bot
   - Register Azure AD application resource in Azure portal
   - In Azure portal, create a [Azure Bot resource](https://docs.microsoft.com/azure/bot-service/bot-builder-authentication?view=azure-bot-service-4.0&tabs=csharp%2Caadv2).

   - Ensure that you've [enabled the Teams Channel](https://docs.microsoft.com/azure/bot-service/channel-connect-teams?view=azure-bot-service-4.0)
   - While registering the bot, use `https://<your_tunnel_domain>/api/messages` as the messaging endpoint.

  **NOTE:** When you create your bot you will create an App ID and App password - make sure you keep these for later.

2) Setup Serveo  
   Run ssh -R 80:localhost:3978 serveo.net

   ```bash
   ssh -R 80:localhost:3978 serveo.net
   ```  

   Alternatively, you can also use the `dev tunnels`. Please follow [Create and host a dev tunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started?tabs=windows) and host the tunnel with anonymous user access command as shown below:

   ```bash
   devtunnel host -p 3978 --allow-anonymous
   ```

3) Setup for code   
- Clone the repository

    ```bash
    git clone https://github.com/OfficeDev/Microsoft-Teams-Samples.git

- Navigate to `samples/meetings-events/csharp` 
    - Modify the `/appsettings.json` and fill in the `{{ MicrosoftAppId }}`,`{{ MicrosoftAppPassword }}` with the values received while doing Microsoft Entra ID app registration in step 1.

- Run the app from a terminal or from Visual Studio, choose option A or B.

  A) From a terminal

  ```bash
  # run the app
  dotnet run
  ```

  B) Or from Visual Studio

  - Launch Visual Studio
  - File -> Open -> Project/Solution
  - Navigate to `MeetingEvents` folder
  - Select `MeetingEvents.csproj` file
  - Press `F5` to run the project

4) Setup Manifest for Teams

Modify the `manifest.json` in the `/AppManifest` folder and replace the following details

   - `<<App-ID>>` with your Microsoft Entra ID app registration id   
   - `<<VALID DOMAIN>>` with base Url domain. E.g. if you are using Serveo it would be `https://1b6xxxx62c253d270b610ec09a7b3b39a17.serveo.net/` then your domain-name will be `1b6xxxx62c253d270b610ec09a7b3b39a17.serveo.net` and if you are using dev tunnels then your domain will be like: `12345.devtunnels.ms`.
   - Zip the contents of `AppManifest` folder into a `manifest.zip`, and use the `manifest.zip` to deploy in app store
   - - **Upload** the `manifest.zip` to Teams
         - Select **Apps** from the left panel.
         - Then select **Upload a custom app** from the lower right corner.
         - Then select the `manifest.zip` file from `AppManifest`.
         - [Install the App in Teams Meeting](https://docs.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/teams-apps-in-meetings?view=msteams-client-js-latest#meeting-lifecycle-scenarios)

**Note**: If you are facing any issue in your app, please uncomment [this](https://github.com/OfficeDev/Microsoft-Teams-Samples/blob/main/samples/meetings-events/csharp/MeetingEvents/AdapterWithErrorHandler.cs#L25) line and put your debugger for local debug.

## Running the sample
Once the meeting where the bot is added starts or ends, real-time updates are posted in the chat.

**MeetingEvents command interaction:**   

![Meeting start event](MeetingEvents/Images/meeting-start.png)

**End meeting events details:**   

![Meeting end event](MeetingEvents/Images/meeting-end.png)

**MeetingParticipantEvents command interaction:**   

To utilize this feature, please enable Meeting event subscriptions for `Participant Join` and `Participant Leave` in your bot, following the guidance outlined in the [meeting participant events](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis?branch=pr-en-us-8455&tabs=channel-meeting%2Cguest-user%2Cone-on-one-call%2Cdotnet%2Cparticipant-join-event#receive-meeting-participant-events) documentation

![Meeting participant added event](MeetingEvents/Images/meeting-participant-added.png)

**End meeting events details:**   

![Meeting participant left event](MeetingEvents/Images/meeting-participant-left.png)

## Deploy the bot to Azure

To learn more about deploying a bot to Azure, see [Deploy your bot to Azure](https://aka.ms/azuredeployment) for a complete list of deployment instructions.

## Further reading

- [Real-time Teams meeting events](https://docs.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/api-references?tabs=dotnet)
- [Meeting apps APIs](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis?tabs=dotnet)
- [Real-time Teams meeting participant events](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/meeting-apps-apis?branch=pr-en-us-8455&tabs=channel-meeting%2Cguest-user%2Cone-on-one-call%2Cdotnet%2Cparticipant-join-event#receive-meeting-participant-events)

<img src="https://pnptelemetry.azurewebsites.net/microsoft-teams-samples/samples/meetings-events-csharp" />