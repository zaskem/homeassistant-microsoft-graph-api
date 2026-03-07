"""Sensor platform for Microsoft Graph API integration."""

from __future__ import annotations

from datetime import timedelta
import logging
from typing import Any

from homeassistant.components.sensor import SensorEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import (
    CoordinatorEntity,
    DataUpdateCoordinator,
    UpdateFailed,
)

from .api import msGraphApiClient
from .const import (
    ATTR_ACCOUNT_ENABLED,
    ATTR_CREATED_DATETIME,
    ATTR_DEVICE_GROUPS,
    ATTR_DEVICE_ID,
    ATTR_DEVICE_OWNERSHIP,
    ATTR_DISPLAY_NAME,
    ATTR_ENROLLMENT_TYPE,
    ATTR_EXTENSION_ATTRIBUTES,
    ATTR_GROUP_ID,
    ATTR_GROUP_MEMBERS,
    ATTR_GROUP_NAME,
    ATTR_GROUP_TYPES,
    ATTR_IS_COMPLIANT,
    ATTR_LAST_SIGNIN,
    ATTR_MANUFACTURER,
    ATTR_MODEL,
    ATTR_OPERATING_SYSTEM,
    ATTR_OS_VERSION,
    ATTR_SECURITY_ENABLED,
    ATTR_TRUST_TYPE,
    ATTR_USER_DEVICES,
    ATTR_USER_ID,
    ATTR_USER_MAIL,
    ATTR_USER_NAME,
    ATTR_USER_PRINCIPAL_NAME,
    ATTR_USER_EMPLOYEE_ID,
    ATTR_USER_JOB_TITLE,
    ATTR_USER_DEPARTMENT,
    DOMAIN,
)

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up Microsoft Graph API sensor platform."""
    client: msGraphApiClient = entry.runtime_data

    # Get privacy mode setting
    privacy_mode = hass.data.get(DOMAIN, {}).get(f"{entry.entry_id}_privacy_mode", True)
    
    # Get safe mode setting
    safe_mode = entry.options.get(
        "safe_mode",
        entry.data.get("safe_mode", True)
    )

    # Create coordinators to manage data updates
    device_coordinator = DeviceCoordinator(hass, client, entry)
    group_coordinator = GroupCoordinator(hass, client)
    user_coordinator = UserCoordinator(hass, client, entry)
    
    # Store coordinators in hass.data for access by other platforms
    hass.data[f"{DOMAIN}_{entry.entry_id}_coordinator"] = device_coordinator
    hass.data[f"{DOMAIN}_{entry.entry_id}_group_coordinator"] = group_coordinator
    hass.data[f"{DOMAIN}_{entry.entry_id}_user_coordinator"] = user_coordinator
    
    # Fetch initial data
    await device_coordinator.async_config_entry_first_refresh()
    await group_coordinator.async_config_entry_first_refresh()
    await user_coordinator.async_config_entry_first_refresh()

    # Add sensor entities
    sensors = [
        GraphAPIDevicesSensor(device_coordinator),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_DEVICE_ID, "Device ID", "mdi:identifier"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_DEVICE_OWNERSHIP, "Device Ownership", "mdi:account-key"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_ENROLLMENT_TYPE, "Enrollment Type", "mdi:application-import"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_IS_COMPLIANT, "Is Compliant", "mdi:check-circle"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_OPERATING_SYSTEM, "Operating System", "mdi:laptop"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_OS_VERSION, "OS Version", "mdi:numeric"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_MANUFACTURER, "Manufacturer", "mdi:factory"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_MODEL, "Model", "mdi:devices"),
        GraphAPIDeviceDetailsSensor(device_coordinator, ATTR_LAST_SIGNIN, "Last Sign In", "mdi:clock-outline"),
        GraphAPIDeviceGroupsSensor(device_coordinator),
        GraphAPIBitLockerKeysSensor(device_coordinator),
    ]
    
    # Only add extension attributes sensor when safe mode is disabled
    if not safe_mode:
        sensors.append(GraphAPIDeviceExtensionAttributesSensor(device_coordinator))
    
    sensors.extend([
        GraphAPIGroupsSensor(group_coordinator),
        GraphAPIGroupDetailsSensor(group_coordinator, ATTR_GROUP_ID, "Group ID", "mdi:identifier"),
        GraphAPIGroupMembersSensor(group_coordinator),
        GraphAPIUsersSensor(user_coordinator),
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_ID, "User ID", "mdi:identifier", False),
    ])
    
    # Add sensitive user sensors (they will show "Hidden" when privacy mode is enabled)
    sensors.extend([
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_MAIL, "User Mail", "mdi:email", True),
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_PRINCIPAL_NAME, "User Principal Name", "mdi:account", True),
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_EMPLOYEE_ID, "User Employee ID", "mdi:badge-account", True),
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_JOB_TITLE, "User Job Title", "mdi:briefcase", True),
        GraphAPIUserDetailsSensor(user_coordinator, ATTR_USER_DEPARTMENT, "User Department", "mdi:office-building", True),
    ])
    
    sensors.append(GraphAPIUserDevicesSensor(user_coordinator))
    
    async_add_entities(sensors, True)


class DeviceCoordinator(DataUpdateCoordinator):
    """Class to manage fetching Graph API device data."""

    def __init__(self, hass: HomeAssistant, client: msGraphApiClient, entry: ConfigEntry) -> None:
        """Initialize the coordinator."""
        super().__init__(
            hass,
            _LOGGER,
            name=DOMAIN,
            update_interval=timedelta(seconds=client.update_interval),
        )
        self.client = client
        self.entry = entry
        self._selected_device: str | None = None

    def _get_privacy_mode(self) -> bool:
        """Get the privacy mode setting."""
        return self.hass.data.get(DOMAIN, {}).get(f"{self.entry.entry_id}_privacy_mode", True)

    def set_selected_device(self, device_name: str) -> None:
        """Set the currently selected device."""
        self._selected_device = device_name

    def get_selected_device(self) -> str | None:
        """Get the currently selected device."""
        return self._selected_device

    async def _async_update_data(self) -> dict[str, Any]:
        """Fetch data from Microsoft Graph API."""
        try:
            devices = await self.client.get_devices()
            
            # Create a dictionary of devices by display name (like PHP script)
            device_dict = {}
            device_names = []
            
            for device in devices:
                display_name = device.get("displayName", "Unknown")
                device_dict[display_name] = device
                device_names.append(display_name)
            
            # If we have a selected device, fetch its groups and BitLocker keys
            device_groups = []
            bitlocker_keys = []
            if self._selected_device and self._selected_device in device_dict:
                selected_device_data = device_dict[self._selected_device]
                device_object_id = selected_device_data.get("id", "")
                device_id = selected_device_data.get("deviceId", "")
                
                # Fetch device groups
                if device_object_id:
                    device_groups = await self.client.get_device_groups(device_object_id)
                
                # Fetch BitLocker recovery keys only if privacy mode is disabled
                if device_id and not self._get_privacy_mode():
                    bitlocker_keys = await self.client.get_bitlocker_recovery_keys(device_id)
                elif device_id and self._get_privacy_mode():
                    bitlocker_keys = ["Hidden (Privacy Mode enabled)"]
            
            return {
                "devices": device_names,
                "device_dict": device_dict,
                "device_count": len(devices),
                "device_groups": device_groups,
                "bitlocker_keys": bitlocker_keys,
                "selected_device": self._selected_device,
            }
        except Exception as err:
            raise UpdateFailed(f"Error fetching devices: {err}") from err


class GroupCoordinator(DataUpdateCoordinator):
    """Class to manage fetching Graph API group data."""

    def __init__(self, hass: HomeAssistant, client: msGraphApiClient) -> None:
        """Initialize the coordinator."""
        super().__init__(
            hass,
            _LOGGER,
            name=f"{DOMAIN}_groups",
            update_interval=timedelta(seconds=client.update_interval),
        )
        self.client = client
        self._selected_group: str | None = None

    def set_selected_group(self, group_name: str) -> None:
        """Set the currently selected group."""
        self._selected_group = group_name

    def get_selected_group(self) -> str | None:
        """Get the currently selected group."""
        return self._selected_group

    async def _async_update_data(self) -> dict[str, Any]:
        """Fetch data from Microsoft Graph API."""
        try:
            groups = await self.client.get_groups()
            
            # Create a dictionary of groups by display name (like PHP script)
            group_dict = {}
            group_names = []
            
            for group in groups:
                display_name = group.get("displayName", "Unknown")
                group_dict[display_name] = group
                group_names.append(display_name)
            
            # If we have a selected group, fetch its members
            group_members = []
            if self._selected_group and self._selected_group in group_dict:
                selected_group_data = group_dict[self._selected_group]
                group_id = selected_group_data.get("id", "")
                
                # Fetch group members (devices)
                if group_id:
                    members_data = await self.client.get_group_members(group_id)
                    # Extract device names from members
                    group_members = [m.get("displayName", "Unknown") for m in members_data]
            
            return {
                "groups": group_names,
                "group_dict": group_dict,
                "group_count": len(groups),
                "group_members": group_members,
                "selected_group": self._selected_group,
            }
        except Exception as err:
            raise UpdateFailed(f"Error fetching groups: {err}") from err


class UserCoordinator(DataUpdateCoordinator):
    """Class to manage fetching Graph API user data."""

    def __init__(self, hass: HomeAssistant, client: msGraphApiClient, entry: ConfigEntry) -> None:
        """Initialize the coordinator."""
        super().__init__(
            hass,
            _LOGGER,
            name=f"{DOMAIN}_users",
            update_interval=timedelta(seconds=client.update_interval),
        )
        self.client = client
        self.entry = entry
        self._selected_user: str | None = None

    def _get_privacy_mode(self) -> bool:
        """Get the privacy mode setting."""
        return self.hass.data.get(DOMAIN, {}).get(f"{self.entry.entry_id}_privacy_mode", True)

    def set_selected_user(self, user_name: str) -> None:
        """Set the currently selected user."""
        self._selected_user = user_name

    def get_selected_user(self) -> str | None:
        """Get the currently selected user."""
        return self._selected_user

    async def _async_update_data(self) -> dict[str, Any]:
        """Fetch data from Microsoft Graph API."""
        try:
            users = await self.client.get_users()
            
            # Create a dictionary of users by display name (like PHP script)
            user_dict = {}
            user_names = []
            
            for user in users:
                display_name = user.get("displayName", "Unknown")
                user_dict[display_name] = user
                user_names.append(display_name)
            
            # If we have a selected user, fetch their owned devices
            user_devices = []
            if self._selected_user and self._selected_user in user_dict:
                selected_user_data = user_dict[self._selected_user]
                user_id = selected_user_data.get("id", "")
                
                # Fetch user owned devices
                if user_id:
                    user_devices = await self.client.get_user_devices(user_id)
            
            return {
                "users": user_names,
                "user_dict": user_dict,
                "user_count": len(users),
                "user_devices": user_devices,
                "selected_user": self._selected_user,
            }
        except Exception as err:
            raise UpdateFailed(f"Error fetching users: {err}") from err


class GraphAPIDevicesSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing list of Graph API devices."""

    def __init__(self, coordinator: DeviceCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Devices"
        self._attr_unique_id = f"{DOMAIN}_devices"
        self._attr_icon = "mdi:devices"

    @property
    def native_value(self) -> int:
        """Return the number of devices."""
        if self.coordinator.data is None:
            return 0
        return self.coordinator.data.get("device_count", 0)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"devices": [], "device_details": {}}
        
        return {
            "devices": self.coordinator.data.get("devices", []),
            "device_details": self.coordinator.data.get("device_dict", {}),
        }


class GraphAPIDeviceDetailsSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing details for the selected Graph API device."""

    def __init__(
        self,
        coordinator: DeviceCoordinator,
        attribute_key: str,
        name: str,
        icon: str,
    ) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attribute_key = attribute_key
        self._attr_name = f"Graph API {name}"
        self._attr_unique_id = f"{DOMAIN}_{attribute_key}"
        self._attr_icon = icon

    @property
    def native_value(self) -> str | None:
        """Return the value of the attribute for the selected device."""
        if self.coordinator.data is None:
            return "Unknown"
        
        selected_device = self.coordinator.data.get("selected_device")
        if not selected_device:
            return "No device selected"
        
        device_dict = self.coordinator.data.get("device_dict", {})
        if selected_device not in device_dict:
            return "Device not found"
        
        device_data = device_dict[selected_device]
        
        # Map attribute keys to device data fields
        field_mapping = {
            ATTR_DEVICE_ID: "deviceId",
            ATTR_DEVICE_OWNERSHIP: "deviceOwnership",
            ATTR_ENROLLMENT_TYPE: "enrollmentType",
            ATTR_IS_COMPLIANT: "isCompliant",
            ATTR_OPERATING_SYSTEM: "operatingSystem",
            ATTR_OS_VERSION: "operatingSystemVersion",
            ATTR_MANUFACTURER: "manufacturer",
            ATTR_MODEL: "model",
            ATTR_LAST_SIGNIN: "approximateLastSignInDateTime",
            ATTR_TRUST_TYPE: "trustType",
            ATTR_ACCOUNT_ENABLED: "accountEnabled",
            ATTR_DISPLAY_NAME: "displayName",
        }
        
        field_name = field_mapping.get(self._attribute_key, self._attribute_key)
        value = device_data.get(field_name)
        
        # Handle None values with appropriate defaults
        if value is None:
            return "Not available"
        
        # Convert boolean to string
        if isinstance(value, bool):
            return "Yes" if value else "No"
        
        return str(value)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {}
        
        selected_device = self.coordinator.data.get("selected_device")
        if not selected_device:
            return {}
        
        return {
            "selected_device": selected_device,
        }


class GraphAPIDeviceGroupsSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing group memberships for the selected Graph API device."""

    def __init__(self, coordinator: DeviceCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Device Groups"
        self._attr_unique_id = f"{DOMAIN}_device_groups"
        self._attr_icon = "mdi:account-group"

    @property
    def native_value(self) -> int:
        """Return the number of groups the device belongs to."""
        if self.coordinator.data is None:
            return 0
        
        groups = self.coordinator.data.get("device_groups", [])
        return len(groups)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"groups": [], "selected_device": None}
        
        return {
            "groups": self.coordinator.data.get("device_groups", []),
            "selected_device": self.coordinator.data.get("selected_device"),
        }


class GraphAPIBitLockerKeysSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing BitLocker recovery keys for the selected Graph API device.
    
    Performs a two-step REST request:
    1. Fetches list of recovery key IDs for the device
    2. Fetches the actual key value for each ID
    
    Note: The 'keys' attribute is excluded from recorder to prevent
    sensitive BitLocker recovery keys from being stored in the database.
    """

    # Exclude sensitive attributes from being recorded in the database
    _unrecorded_attributes = frozenset({"keys"})

    def __init__(self, coordinator: DeviceCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API BitLocker Recovery Keys"
        self._attr_unique_id = f"{DOMAIN}_bitlocker_keys"
        self._attr_icon = "mdi:lock-reset"

    @property
    def native_value(self) -> int | str:
        """Return the number of recovery keys available."""
        if self.coordinator.data is None:
            return 0
        
        keys = self.coordinator.data.get("bitlocker_keys", [])
        
        # Handle status messages
        if keys and isinstance(keys[0], str):
            status_messages = [
                "No keys available",
                "Authentication failed",
                "Token expired",
                "Failed to fetch recovery key list",
                "Keys found but could not retrieve values",
                "Network error",
                "Unexpected error",
            ]
            if keys[0] in status_messages:
                return keys[0]
        
        return len(keys)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {
                "keys": [],
                "selected_device": None,
                "status": "No data",
            }
        
        keys = self.coordinator.data.get("bitlocker_keys", [])
        selected_device = self.coordinator.data.get("selected_device")
        
        # Determine status
        status = "Ready"
        if not selected_device:
            status = "No device selected"
        elif not keys:
            status = "No keys available"
        elif isinstance(keys[0], str) and keys[0] in [
            "Authentication failed",
            "Token expired",
            "Failed to fetch recovery key list",
            "Network error",
            "Unexpected error",
        ]:
            status = keys[0]
        else:
            status = f"{len(keys)} key(s) available"
        
        return {
            "keys": keys,
            "selected_device": selected_device,
            "status": status,
        }


class GraphAPIDeviceExtensionAttributesSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing extension attributes for the selected Graph API device.
    
    Note: This sensor is only available when Safe Mode is disabled.
    Extension attributes are custom fields that can store additional device metadata.
    """

    def __init__(self, coordinator: DeviceCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Device Extension Attributes"
        self._attr_unique_id = f"{DOMAIN}_device_extension_attributes"
        self._attr_icon = "mdi:tag-multiple"

    @property
    def native_value(self) -> int | str:
        """Return the number of non-null extension attributes."""
        if self.coordinator.data is None:
            return "Unknown"
        
        selected_device = self.coordinator.data.get("selected_device")
        if not selected_device:
            return "No device selected"
        
        device_dict = self.coordinator.data.get("device_dict", {})
        if selected_device not in device_dict:
            return "Device not found"
        
        device_data = device_dict[selected_device]
        extension_attrs = device_data.get("extensionAttributes", {})
        
        # Count non-null attributes
        non_null_count = sum(1 for v in extension_attrs.values() if v is not None)
        
        if non_null_count == 0:
            return "No attributes set"
        
        return f"{non_null_count} attribute(s) set"

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return extension attributes."""
        if self.coordinator.data is None:
            return {}
        
        selected_device = self.coordinator.data.get("selected_device")
        if not selected_device:
            return {"selected_device": None}
        
        device_dict = self.coordinator.data.get("device_dict", {})
        if selected_device not in device_dict:
            return {"selected_device": selected_device}
        
        device_data = device_dict[selected_device]
        extension_attrs = device_data.get("extensionAttributes", {})
        
        # Return all extension attributes, including nulls
        return {
            "selected_device": selected_device,
            "extension_attributes": extension_attrs,
            # Also flatten them for easier access
            **{f"extensionAttribute{i}": extension_attrs.get(f"extensionAttribute{i}") 
               for i in range(1, 16)},
        }


class GraphAPIGroupsSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing list of Graph API security groups."""

    def __init__(self, coordinator: GroupCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Groups"
        self._attr_unique_id = f"{DOMAIN}_groups"
        self._attr_icon = "mdi:account-group"

    @property
    def native_value(self) -> int:
        """Return the number of groups."""
        if self.coordinator.data is None:
            return 0
        return self.coordinator.data.get("group_count", 0)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"groups": [], "group_details": {}}
        
        return {
            "groups": self.coordinator.data.get("groups", []),
            "group_details": self.coordinator.data.get("group_dict", {}),
        }


class GraphAPIGroupDetailsSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing details for the selected Graph API group."""

    def __init__(
        self,
        coordinator: GroupCoordinator,
        attribute_key: str,
        name: str,
        icon: str,
    ) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attribute_key = attribute_key
        self._attr_name = f"Graph API {name}"
        self._attr_unique_id = f"{DOMAIN}_{attribute_key}"
        self._attr_icon = icon

    @property
    def native_value(self) -> str | None:
        """Return the value of the attribute for the selected group."""
        if self.coordinator.data is None:
            return "Unknown"
        
        selected_group = self.coordinator.data.get("selected_group")
        if not selected_group:
            return "No group selected"
        
        group_dict = self.coordinator.data.get("group_dict", {})
        if selected_group not in group_dict:
            return "Group not found"
        
        group_data = group_dict[selected_group]
        
        # Map attribute keys to group data fields
        field_mapping = {
            ATTR_GROUP_ID: "id",
            ATTR_GROUP_NAME: "displayName",
            ATTR_SECURITY_ENABLED: "securityEnabled",
            ATTR_GROUP_TYPES: "groupTypes",
            ATTR_CREATED_DATETIME: "createdDateTime",
        }
        
        field_name = field_mapping.get(self._attribute_key, self._attribute_key)
        value = group_data.get(field_name)
        
        # Handle None values with appropriate defaults
        if value is None:
            return "Not available"
        
        # Convert boolean to string
        if isinstance(value, bool):
            return "Yes" if value else "No"
        
        # Convert lists to comma-separated strings
        if isinstance(value, list):
            return ", ".join(str(v) for v in value) if value else "None"
        
        return str(value)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {}
        
        selected_group = self.coordinator.data.get("selected_group")
        if not selected_group:
            return {}
        
        return {
            "selected_group": selected_group,
        }


class GraphAPIGroupMembersSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing device members of the selected Graph API group."""

    def __init__(self, coordinator: GroupCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Group Members"
        self._attr_unique_id = f"{DOMAIN}_group_members"
        self._attr_icon = "mdi:server-network"

    @property
    def native_value(self) -> int:
        """Return the number of device members in the group."""
        if self.coordinator.data is None:
            return 0
        
        members = self.coordinator.data.get("group_members", [])
        return len(members)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"members": [], "selected_group": None}
        
        return {
            "members": self.coordinator.data.get("group_members", []),
            "selected_group": self.coordinator.data.get("selected_group"),
        }


class GraphAPIUsersSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing list of Graph API users."""

    def __init__(self, coordinator: UserCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Users"
        self._attr_unique_id = f"{DOMAIN}_users"
        self._attr_icon = "mdi:account-multiple"

    @property
    def native_value(self) -> int:
        """Return the number of users."""
        if self.coordinator.data is None:
            return 0
        return self.coordinator.data.get("user_count", 0)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"users": [], "user_details": {}}
        
        return {
            "users": self.coordinator.data.get("users", []),
            "user_details": self.coordinator.data.get("user_dict", {}),
        }


class GraphAPIUserDetailsSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing details for the selected Graph API user."""

    def __init__(
        self,
        coordinator: UserCoordinator,
        attribute_key: str,
        name: str,
        icon: str,
        privacy_sensitive: bool = False,
    ) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attribute_key = attribute_key
        self._privacy_sensitive = privacy_sensitive
        self._attr_name = f"Graph API {name}"
        self._attr_unique_id = f"{DOMAIN}_{attribute_key}"
        self._attr_icon = icon

    @property
    def native_value(self) -> str | None:
        """Return the value of the attribute for the selected user."""
        # Check if this is a privacy-sensitive field and privacy mode is enabled
        if self._privacy_sensitive and self.coordinator._get_privacy_mode():
            return "Hidden (Privacy Mode enabled)"
        
        if self.coordinator.data is None:
            return "Unknown"
        
        selected_user = self.coordinator.data.get("selected_user")
        if not selected_user:
            return "No user selected"
        
        user_dict = self.coordinator.data.get("user_dict", {})
        if selected_user not in user_dict:
            return "User not found"
        
        user_data = user_dict[selected_user]
        
        # Map attribute keys to user data fields
        field_mapping = {
            ATTR_USER_ID: "id",
            ATTR_USER_NAME: "displayName",
            ATTR_USER_MAIL: "mail",
            ATTR_USER_PRINCIPAL_NAME: "userPrincipalName",
            ATTR_USER_EMPLOYEE_ID: "employeeId",
            ATTR_USER_JOB_TITLE: "jobTitle",
            ATTR_USER_DEPARTMENT: "department",
        }
        
        field_name = field_mapping.get(self._attribute_key, self._attribute_key)
        value = user_data.get(field_name)
        
        # Handle None values with appropriate defaults
        if value is None:
            return "Not available"
        
        return str(value)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {}
        
        selected_user = self.coordinator.data.get("selected_user")
        if not selected_user:
            return {}
        
        return {
            "selected_user": selected_user,
        }


class GraphAPIUserDevicesSensor(CoordinatorEntity, SensorEntity):
    """Sensor showing devices owned by the selected Graph API user."""

    def __init__(self, coordinator: UserCoordinator) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_name = "Graph API User Devices"
        self._attr_unique_id = f"{DOMAIN}_user_devices"
        self._attr_icon = "mdi:devices"

    @property
    def native_value(self) -> int:
        """Return the number of devices owned by the user."""
        if self.coordinator.data is None:
            return 0
        
        devices = self.coordinator.data.get("user_devices", [])
        return len(devices)

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        if self.coordinator.data is None:
            return {"devices": [], "selected_user": None}
        
        return {
            "devices": self.coordinator.data.get("user_devices", []),
            "selected_user": self.coordinator.data.get("selected_user"),
        }



