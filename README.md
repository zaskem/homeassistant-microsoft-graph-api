# Microsoft Graph API Integration for Home Assistant

This integration allows you to work with devices, users, and groups from a Microsoft Tenant (Entra ID, Intune) in Home Assistant by connecting to the Microsoft Graph API.

**IMPORTANT NOTE**: This integration is intended for educational/investigative purposes only and should _NOT_ be used in a production environment unless you fully understand the implications or wish to run the risk of a Resume-Generating Event (RGE). The integration as provided will not generally cause irreparable harm; however, you _are_ connecting to a directory/tenant with potential write-level access and actions.

## Features

- Authenticate with Microsoft Graph API using an Azure application registration
- Fetch and display a list of devices, users, and groups in Entra ID/Intune
- Automatic token refresh using client secret or certificate authentication
- Automatic polling of data
- Services provided to enable write-back to Entra ID data points

## Prerequisites

Before setting up this integration, you need to create an Azure application registration with the appropriate permissions:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** > **App registrations**
3. Click **New registration**
4. Give your application a name (e.g., "Home Assistant Graph API Sandbox")
5. Select **Single tenant only**
6. Click **Register**
7. Note down the **Application (client) ID** and **Directory (tenant) ID**
8. Go to **Manage** > **Certificates & secrets** > **New client secret**
   - This integration supports certificate-based authentication as well; see the [Certificate Authentiation section below](#using-certificate-based-authentication) for additional information.
9.  Create a new secret and note down its **Value** (this is your client secret)
10. Go to **API permissions** > **Add a permission**
11. Select **Microsoft Graph** > **Application permissions**
12. Add the following initial permissions:
    - `Device.Read.All`
    - `Directory.Read.All`
    - `User.Read.All`
  - This is not a [comprehensive list of permissions](#comprehensive-list-of-app-registration-permissions-for-integration) but are enough to get started. You need to add permissions for all sensors to behave (see [addendum below for comprehensive list](#comprehensive-list-of-app-registration-permissions-for-integration)).
13. Click **Grant admin consent** for your organization

## Installation

1. Copy the `ha_ms_graph_api` folder to your Home Assistant `custom_components` directory
2. Restart Home Assistant
3. Go to **Settings** > **Devices & Services**
4. Click **Add Integration**
5. Search for "Microsoft Graph API Sandbox"
6. Enter your Azure Application credentials:
   - **Azure Application Client ID**: The Application (client) ID from step 7 above
   - **Azure Application Client Secret**: The secret value from step 9 above
   - **Azure Tenant ID**: The Directory (tenant) ID from step 7 above
7. Select optional settings:
   - **Use Certificate-Based Authentication (optional)**: To use host-based [certificate authentication](#using-certificate-based-authentication) (substantially better for restricting access to a Home Assistant instance)
   - **Azure Application Certificate Path (optional)**: The local path to your certificate private key
   - **Update Interval (seconds)**: The polling interval for refresh of sensors (default 300 -- five minutes)
   - **Safe Mode (read-only actions)**: Checked (default) disables write-action services. Unchecked allows write-action services _and requires additional permissions (see sections below)_
   - **Privacy Mode (hide sensitive data)**: Checked (default) hides or makes unavailable certain sensors in the integration (including Bitlocker keys and some user properties). Unchecked enables all sensors.
8. Click **Submit**

## Configuration Options

After setup, you can reconfigure the following options by going to **Settings** > **Devices & Services** > **Microsoft Graph API Sandbox** > **Configure**:

- **Azure Application Client Secret**: Rotating secrets is a good thing; this setting allows you to change a secret without completely re-installing the integration
- **Use Certificate-Based Authentication**: Switch between authentication methods
- **Azure Application Certificate Path**: Local path to combined certificate and private key if Certificate-Based Authentication is used
- **Update Interval**: How often to poll the Microsoft Graph API (default: 300 seconds)
- **Safe Mode**: When enabled (default), prevents write operations to Entra ID/Intune
- **Privacy Mode**: When enabled (default), displays "Hidden (Privacy Mode enabled)" for BitLocker keys and sensitive user data (email, principal name, employee ID, job title, department) instead of the actual values

## Sensor List

Following setup, the integration will create or expose a number of sensors (listed in alphabetical order):
* Graph API BitLocker Recovery Keys
* Graph API Device Extension Attributes (plus individual selector and editor)
* Graph API Device Groups
* Graph API Device ID
* Graph API Device Ownership
* Graph API Device Selector
* Graph API Devices
* Graph API Enrollment Type
* Graph API Group ID
* Graph API Group Members
* Graph API Group Selector
* Graph API Groups
* Graph API Is Compliant
* Graph API Last Sign In
* Graph API Manufacturer
* Graph API Model
* Graph API Operating System
* Graph API OS Version
* Graph API Privacy Mode (read-only; matches configuration setting)
* Graph API Safe Mode (read-only; matches configuration setting)
* Graph API User Department (plus editor)
* Graph API User Devices
* Graph API User Employee ID (plus editor)
* Graph API User ID
* Graph API User Job Title (plus editor)
* Graph API User Mail
* Graph API User Principal Name
* Graph API User Selector
* Graph API Users

Sensors will update automatically as actions are called or polled. Note a short refresh delay of 1-10 seconds is normal, especially when frequent calls are made such as selecting different devices or users in short order.

### Privacy and Safe Mode Sensor Behavior

When Privacy Mode is enabled (default), the following sensors will display "Hidden (Privacy Mode enabled)" instead of actual values:
- Graph API BitLocker Recovery Keys
- Graph API User Mail
- Graph API User Principal Name
- Graph API User Employee ID
- Graph API User Job Title
- Graph API User Department

When Safe Mode is enabled (default), the following sensors and editable field sensors will not be provided by the integration:
- Graph API Device Extension Attributes
- Graph API Device Extension Attribute Editor
- Graph API User Department Editor
- Graph API User Employee ID Editor
- Graph API User Job Title Editor

### Special/Convenience Sensors

#### Graph API Privacy Mode and Safe Mode Sensors

The Privacy Mode and Safe Mode sensors mirror the current configuration. To change either sensor, reconfigure the integration through **Settings** > **Devices & Services** > **Microsoft Graph API Sandbox** > **Configure**.

#### Graph API Extension Attribute Selector

A special dropdown selector is provided to choose a device extension attribute (1-15) to assist calling the `update_device_extension_attribute` service. This makes it easier to change attributes without creating a helper dropdown.

#### Graph API Extension Attribute and User Property Editors

A set of editable text fields that display and afford editing of the selected extension attribute value or user property values. These sensors automatically sync with:
- The currently selected device and selected extension attribute (from Graph API Extension Attribute Selector)
- The currently selected user (from Graph API User Selector)


## Pre-Configured Dashboard, Automations, and Examples

Descriptions and information to install a pre-configured Home Assistant dashboard with basic actions and other automation examples can be [found in `examples.md`](examples.md). 

## Using Certificate-Based Authentication

In an integration like this, using host-specific certificate-based authentication is advised over client secrets. Certificate-based authentication more easily restricts or identifies/affords access to specific hosts. You can create and use a self-signed certificate from your Home Assistant instance.

To configure and set up Certificate authentication:

1. On your Home Assistant instance, use an SSH method to access the system, and generate a certificate and key file:
   - `cd config`
   - `mkdir ssl`
   - `cd ssl`
   - `openssl req -new -x509 -days 365 -nodes -keyout ha_graphapi.key -out ha_graphapi.crt`
   - Follow the prompts (default values are acceptable)
2. Combine the key and certificate into one file (referenced by Home Assistant):
   - `cat ha_graphapi.key ha_graphapi.crt > ha_graphapi.pem`
3. Make a local copy of the `ha_graphapi.crt` file (or its contents) to upload to Azure
4. In the Azure application registration, go to **Certificates & secrets** > **Upload certificate** and provide the `ha_graphapi.crt` file and give it a reasonable description
5. In Home Assistant, navigate to **Settings** > **Devices & Services** > **Microsoft Graph API Sandbox** > **Configure** and set the following values:
   - **Use Certificate-Based Authentication (optional)**: Check this box
   - **Azure Application Certificate Path (optional)**: Enter the path to the `pem` file from step 2 above (e.g. `/config/ssl/ha_grapapi.pem`)
6. Click **Submit**
7. The integration should automatically reload, but if not use the three dots button > **Reload** and check for errors.

You are now using Certificate-based authentication for the integration. Keeping the secret in place is acceptable, but the secret can be overwritten by setting it to a single space and saved, which will remove any reference to the intial secret.

## Comprehensive List of App Registration Permissions for Integration

Permissions are listed below in alphabetical order by function of the integration. All permissions are Application based and require admin consent for the organization.

### Basic Read-Only Functionality

* Device.Read.All
* Directory.Read.All
* User.Read.All

### Privacy-Mode Disabled (show sensitive data)

* BitlockerKey.Read.All
* BitlockerKey.ReadBasic.All

### Safe Mode Disabled (allow write-back)

* Device.ReadWrite.All
* User.ReadWrite.All (or Directory.ReadWrite.All)

## Removing or Re-Installing This Integration

The integration is quite well-contained. Removing/Deleting it from the **Settings** > **Devices & Services** > **Microsoft Graph API Sandbox** interface is the recommended method and will automatically remove and clean up created sensors, configuration, etc.

The integration can be re-installed by deleting it, then [following the installation instructions](#installation).

**Note:** The example dashboard and automations, if added, will not be removed by deleting the integration and must be manually removed if desired (not re-installing the integration). If re-installing the integration, the dashboard and automations will continue to behave after installation is complete.

To completely eliminate the integration from your Home Assistant instance, remove the `ha_ms_graph_api` folder from the `custom_components` directory after deleting the integration from the UI.

## Troubleshooting

If you encounter authentication issues:

1. Verify your Client ID, Client Secret, and Tenant ID are correct
2. If using Certificate authentication, verify the certificate is properly registered in the Azure application configuration, the Home Asisstant key file includes the public certificate, and the local path to the key file is correct
3. Ensure required API permissions have been added (read and/or write actions)
4. Ensure admin consent has been granted for required API permissions
5. Check Home Assistant logs for detailed error messages

## Credits

This integration is based on an idea @zaskem had built out using Home Assistant's native RESTFul platform by proxying Graph API calls through a third-party host to bring device monitoring and expose device details in Home Assistant. The integration eliminates the need to set up a proxy host (and scripts) to translate/transform requests from Home Assistant to Graph API and is generally a novel idea not intended to be used in a production setting.
