
# SPO.Webhook

This project provides a webhook endpoint that listens for events occurring on a SharePoint list. When an event (such as an item being added, updated, or deleted) occurs, SharePoint calls this API endpoint. However, the webhook notification does not include detailed information about which item was changed. To retrieve specific details about the changes, you must use the SharePoint ChangeQuery API to fetch the changes on that list.

## Usage

Run the application and expose the port using below command. Make sure that ngrok is installed:

```cmd
ngrok http --host-header=localhost 5111  
```


To test this API, you can expose it locally using [ngrok](https://ngrok.com/) and then add a webhook subscription to your SharePoint list using the following PowerShell command:

```powershell
Add-PnPWebhookSubscription -List rer -NotificationUrl https://4353c832ede0.ngrok-free.app/api/spwebhook/handlerequest -ExpirationDate "2025-09-01" -ClientState "A0A354EC-97D4-4D83-9DDB-144077ADB449"
```

To remove the Webhook, use belo command:

```powershell
 Remove-PnPWebhookSubscription -List rer -Identity <subscription-id>
```


> **Note:** The `ClientState` value must match the value that is configured in your webhook application. This is used to validate incoming notifications and ensure they are from a trusted source.

Example output:

```
Id                 : 5cdbb94e-bf3e-4050-a514-1bce60f65aa2
ClientState        : A0A354EC-97D4-4D83-9DDB-144077ADB449
ExpirationDateTime : 01-09-2025 07:00:00
NotificationUrl    : https://4353c832ede0.ngrok-free.app/api/spwebhook/handlerequest
Resource           : a0922814-a5ae-4be6-b16c-754493a8e61e
```
