"""Select platform for Microsoft Graph API integration."""

from __future__ import annotations

import logging
from typing import Any

from homeassistant.components.select import SelectEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from .const import DOMAIN
from .sensor import DeviceCoordinator, GroupCoordinator, UserCoordinator

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up the select sensor platform."""
    # Get the coordinators from the sensor platform
    device_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_coordinator")
    group_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_group_coordinator")
    user_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_user_coordinator")
    
    selects = []
    if device_coordinator:
        selects.append(DeviceSelect(device_coordinator))
    if group_coordinator:
        selects.append(GroupSelect(group_coordinator))
    if user_coordinator:
        selects.append(UserSelect(user_coordinator))
    
    # Add the extension attribute selector (doesn't require a coordinator)
    selects.append(ExtensionAttributeSelect(entry))
    
    if selects:
        async_add_entities(selects, True)


class DeviceSelect(CoordinatorEntity, SelectEntity):
    """Select entity for choosing device."""

    def __init__(self, coordinator: DeviceCoordinator) -> None:
        """Initialize the select entity."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Device Selector"
        self._attr_unique_id = f"{DOMAIN}_device_select"
        self._attr_icon = "mdi:menu"

    @property
    def options(self) -> list[str]:
        """Return the list of available device names."""
        if self.coordinator.data is None:
            return ["No devices available"]
        devices = self.coordinator.data.get("devices", [])
        if devices:
            return ["Select Device"] + devices
        return ["No devices available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected device."""
        selected = self.coordinator.get_selected_device()
        
        # If nothing selected, return the placeholder
        if selected is None:
            if self.options and self.options[0] == "Select Device":
                return "Select Device"
            return self.options[0] if self.options else "No devices available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No devices available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected device."""
        if option in self.options and option not in ["No devices available", "Select Device"]:
            self.coordinator.set_selected_device(option)
            # Trigger coordinator update to fetch groups for new device
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()
        elif option == "Select Device":
            # Clear selection when placeholder is selected
            self.coordinator.set_selected_device(None)
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected device."""
        current = self.current_option
        if current and current not in ["No devices available", "Select Device"] and self.coordinator.data is not None:
            device_dict = self.coordinator.data.get("device_dict", {})
            selected_device = device_dict.get(current, {})
            if selected_device:
                return {
                    "device_id": selected_device.get("deviceId", "Unknown"),
                    "display_name": selected_device.get("displayName", "Unknown"),
                    "operating_system": selected_device.get("operatingSystem", "Unknown"),
                    "operating_system_version": selected_device.get("operatingSystemVersion", "Unknown"),
                    "trust_type": selected_device.get("trustType", "Unknown"),
                    "account_enabled": selected_device.get("accountEnabled", False),
                }
        return {}


class GroupSelect(CoordinatorEntity, SelectEntity):
    """Select entity for choosing a group."""

    def __init__(self, coordinator: GroupCoordinator) -> None:
        """Initialize the select entity."""
        super().__init__(coordinator)
        self._attr_name = "Graph API Group Selector"
        self._attr_unique_id = f"{DOMAIN}_group_select"
        self._attr_icon = "mdi:account-group"

    @property
    def options(self) -> list[str]:
        """Return the list of available group names."""
        if self.coordinator.data is None:
            return ["No groups available"]
        groups = self.coordinator.data.get("groups", [])
        if groups:
            return ["Select Group"] + groups
        return ["No groups available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected group."""
        selected = self.coordinator.get_selected_group()
        
        # If nothing selected, return the placeholder
        if selected is None:
            if self.options and self.options[0] == "Select Group":
                return "Select Group"
            return self.options[0] if self.options else "No groups available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No groups available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected group."""
        if option in self.options and option not in ["No groups available", "Select Group"]:
            self.coordinator.set_selected_group(option)
            # Trigger coordinator update to fetch members for new group
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()
        elif option == "Select Group":
            # Clear selection when placeholder is selected
            self.coordinator.set_selected_group(None)
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected group."""
        current = self.current_option
        if current and current not in ["No groups available", "Select Group"] and self.coordinator.data is not None:
            group_dict = self.coordinator.data.get("group_dict", {})
            selected_group = group_dict.get(current, {})
            if selected_group:
                return {
                    "group_id": selected_group.get("id", "Unknown"),
                    "display_name": selected_group.get("displayName", "Unknown"),
                    "security_enabled": selected_group.get("securityEnabled", False),
                    "group_types": selected_group.get("groupTypes", []),
                    "created_datetime": selected_group.get("createdDateTime", "Unknown"),
                }
        return {}


class UserSelect(CoordinatorEntity, SelectEntity):
    """Select entity for choosing a user."""

    def __init__(self, coordinator: UserCoordinator) -> None:
        """Initialize the select entity."""
        super().__init__(coordinator)
        self._attr_name = "Graph API User Selector"
        self._attr_unique_id = f"{DOMAIN}_user_select"
        self._attr_icon = "mdi:account"

    @property
    def options(self) -> list[str]:
        """Return the list of available user names."""
        if self.coordinator.data is None:
            return ["No users available"]
        users = self.coordinator.data.get("users", [])
        if users:
            return ["Select User"] + users
        return ["No users available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected user."""
        selected = self.coordinator.get_selected_user()
        
        # If nothing selected, return the placeholder
        if selected is None:
            if self.options and self.options[0] == "Select User":
                return "Select User"
            return self.options[0] if self.options else "No users available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No users available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected user."""
        if option in self.options and option not in ["No users available", "Select User"]:
            self.coordinator.set_selected_user(option)
            # Trigger coordinator update to fetch devices for new user
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()
        elif option == "Select User":
            # Clear selection when placeholder is selected
            self.coordinator.set_selected_user(None)
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected user."""
        current = self.current_option
        if current and current not in ["No users available", "Select User"] and self.coordinator.data is not None:
            user_dict = self.coordinator.data.get("user_dict", {})
            selected_user = user_dict.get(current, {})
            if selected_user:
                return {
                    "user_id": selected_user.get("id", "Unknown"),
                    "display_name": selected_user.get("displayName", "Unknown"),
                    "mail": selected_user.get("mail", "Unknown"),
                    "user_principal_name": selected_user.get("userPrincipalName", "Unknown"),
                }
        return {}


class ExtensionAttributeSelect(SelectEntity):
    """Select entity for choosing extension attribute number."""

    # Mapping of user-friendly labels to numeric values
    EXTENSION_ATTRIBUTES = {
        "Extension Attribute 1": 1,
        "Extension Attribute 2": 2,
        "Extension Attribute 3": 3,
        "Extension Attribute 4": 4,
        "Extension Attribute 5": 5,
        "Extension Attribute 6": 6,
        "Extension Attribute 7": 7,
        "Extension Attribute 8": 8,
        "Extension Attribute 9": 9,
        "Extension Attribute 10": 10,
        "Extension Attribute 11": 11,
        "Extension Attribute 12": 12,
        "Extension Attribute 13": 13,
        "Extension Attribute 14": 14,
        "Extension Attribute 15": 15,
    }

    def __init__(self, entry: ConfigEntry) -> None:
        """Initialize the select entity."""
        self._attr_name = "Graph API Extension Attribute Selector"
        self._attr_unique_id = f"{DOMAIN}_{entry.entry_id}_extension_attr_select"
        self._attr_icon = "mdi:numeric"
        self._current_option = "Extension Attribute 1"

    @property
    def options(self) -> list[str]:
        """Return the list of available extension attributes."""
        return list(self.EXTENSION_ATTRIBUTES.keys())

    @property
    def current_option(self) -> str:
        """Return the currently selected extension attribute."""
        return self._current_option

    async def async_select_option(self, option: str) -> None:
        """Change the selected extension attribute."""
        if option in self.EXTENSION_ATTRIBUTES:
            self._current_option = option
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return the numeric value of the selected attribute."""
        return {
            "attribute_number": self.EXTENSION_ATTRIBUTES[self._current_option],
            "description": f"Use this value ({self.EXTENSION_ATTRIBUTES[self._current_option]}) in the update_device_extension_attribute service",
        }
