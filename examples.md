# Examples for Microsoft Graph API Integration

A number of examples of ways to interact with the [Microsoft Graph API Integration for Home Assistant](README.md) are outlined in this documentation. It is broken down in to two key sections:

1. [Pre-Configured Installable Examples](#pre-configured-installable-examples)
2. [Service Documentation and YAML-based Snippets](#service-documentation-and-yaml-based-snippets)

To get started quickly, use the pre-configured examples provided in the repo!

## Pre-Configured Installable Examples

As an example integration, the concept is to get you started quickly so you can play around with Microsoft Graph API via Python and Home Assistant as the interface. This section includes all the necessities to both use Graph API in read-only and write-enabled modes.

### Getting Started | Read-Only Actions
For basic read-only interaction with Graph API all you need to do to start is set up the [provided dashboard](example-yaml/dashboard.yml). Follow these instructions to get your dashboard loaded:

1. Go to **Settings** > **Dashboards** > **+ Add Dashboard** > **New dashboard from scratch**
2. Provide a Title for your dashboard (`Graph API Dashboard` is fine), add an Icon, and change the auto-filled URL if desired
3. Change the dashboard's visibility for Admin only and add it to the sidebar based on your personal preferences
4. Click **Create**
5. Select your newly created dashboard and use the **Edit** (pencil) icon in the upper right-hand corner to put the dashboard in edit mode
6. Use the **three vertical dots menu** in the upper right-hand corner and select **Raw configuration editor**
7. Copy and paste the entire contents of the [`example-yaml/dashboard.yml`](example-yaml/dashboard.yml) into the configuration, overwriting the existing contents
8. Click **Save**
9. Click the **X** option in the header to close the editor
10. Click **Done** in the upper right-hand corner

You will now have an interactive dashboard with two tabs (one for **Read-Only Actions**, one for **Write Actions**) and a three column layout on the Read-Only Actions tab where interactions are grouped into three categories:

* Graph API Groups
* Graph API Devices
* Graph API Users

Congratulations! You can now interact with the Microsoft Graph API from Home Assistant!

### Advanced | Write Actions

The [provided dashboard](example-yaml/dashboard.yml) also creates a basic interface to make write actions using Graph API, available on the **Write Actions** tab where interactions are grouped into two categories:

* Graph API Devices
* Graph API Users

Out of the box with Safe Mode and Privacy Mode enabled (the default), little can be interacted with on this dashboard tab. Before you disable either mode, however, take a few minutes to set up example automations. Seven automations are included for you to use:

* [Example Write User Properties](example-yaml/write-user.yml)
* [Example Write Extension Attribute](example-yaml/write-extensionattribute.yml)
* [Example Refresh User Employee ID Textbox](example-yaml/refresh-user-employee-id-textbox.yml)
* [Example Refresh User Job Title Textbox](example-yaml/refresh-user-job-title-textbox.yml)
* [Example Refresh User Department Textbox](example-yaml/refresh-user-dept-textbox.yml)
* [Example Refresh Extension Attribute Textbox](example-yaml/refresh-extension-attribute-input-textbox.yml)

#### Installing Provided Automations

For each of the provided examples listed above:

1. Go to **Settings** > **Automations & scenes** > **+ Create automation** > **Create new automation (from scratch)**
2. Use the **three vertical dots menu** in the upper right-hand corner and select **Edit in YAML**
3. Copy and paste the entire contents of the provided YAML into the configuration, overwriting the existing contents
4. Click **Save**
5. Click **Rename**

#### Enabling Write Access and Using the Write Actions Dashboard Tab

**Warning:** Proceeding from this point _will make write actions_ if API permissions are enabled. Use with caution.

**API Permissions Requirement:** See the [comprehensive list of permissions](#comprehensive-list-of-app-registration-permissions-for-integration) for necessary API permissions that must be granted to your application registration before write actions will succeed.

When/If you are ready to try write actions, follow these steps to enable the functionality via the Write Actions dashboard tab:

1. Go to **Settings** > **Devices & Services** > **Microsoft Graph API Sandbox** > **Configure**
2. Uncheck **Safe Mode** to expose and change device Extension Attributes
3. Also uncheck **Privacy Mode** to expose and change selected user properties (Employee ID, Job Title, Description)
4. Click **Submit** and **Finish** to reload the integration

The **Write Actions** tab of the dashboard will have additional actions available to edit device Extension Attributes and User Properties (if Privacy Mode was disabled).

#### Writing Device Extension Attributes

To update a device's Extension Attributes via the dashboard:

1. Select a device with the dropdown -- a list of current Extension Attributes for the device is displayed
2. Select the Extension Attribute to update with the associated dropdown
3. The text of the Extension Attribute Editor box will populate with the current data; update this text as necessary
4. Click the adjacent **Update** button

In a few seconds, the device's list of Extension Attributes will update with the updated value!

#### Writing User Properties

To update a user's properties via the dashboard:

1. Select a user with the dropdown -- current values for the user will be populated in the Writable Fields box and the text fields for the User Property Editor box
2. Edit the text fields in the User Property Editor box as necessary
3. Click the adjacent **Update** button

In a few seconds, the user's writable fields will update accordingly!

**Note:** The pre-configured dashboard requires all three user property fields to be updated simultaneously. These values cannot be cleared (set to null/blank) via the dashboard.

It is possible to clear or set individual user properties via this integration, but such an action needs to be configured manually (see YAML section below)

## Service Documentation and YAML-based Snippets

### Update Device Extension Attribute

Service allows you to update extension attributes on devices in Entra ID. **This service is only available when Safe Mode is disabled.**

**Service:** `ha_ms_graph_api.update_device_extension_attribute`

**Parameters:**
- `device_name` (required): The display name of the device to update
- `attribute_number` (required): The extension attribute number (1-15)
- `value` (optional): The value to set. Leave empty or use `null` to clear the attribute

**Example:**
```yaml
service: ha_ms_graph_api.update_device_extension_attribute
data:
  device_name: "LAPTOP-ABC123"
  attribute_number: 1
  value: "IT-Department"
```

**To clear an attribute:**
```yaml
service: ha_ms_graph_api.update_device_extension_attribute
data:
  device_name: "LAPTOP-ABC123"
  attribute_number: 1
  value: null
```

### Update User Properties

Service allows you to update user properties (employee ID, job title, and department) in Entra ID. **This service is only available when Safe Mode is disabled.**

At least one property must be provided in the service call. Properties not specified remain unchanged.

**Service:** `ha_ms_graph_api.update_user_properties`

**Parameters:**
- `user_name` (required): The display name of the user to update
- `employee_id` (optional): The employee ID to set
- `job_title` (optional): The job title to set
- `department` (optional): The department to set

**Example - Update all properties:**
```yaml
service: ha_ms_graph_api.update_user_properties
data:
  user_name: "John Doe"
  employee_id: "EMP-12345"
  job_title: "Senior Engineer"
  department: "Engineering"
```

**Example - Update only job title:**
```yaml
service: ha_ms_graph_api.update_user_properties
data:
  user_name: "John Doe"
  job_title: "Lead Engineer"
```

**Example - Clear a property (set to empty string):**
```yaml
service: ha_ms_graph_api.update_user_properties
data:
  user_name: "John Doe"
  employee_id: ""
```
