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
        return devices if devices else ["No devices available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected device."""
        selected = self.coordinator.get_selected_device()
        
        # If nothing selected and we have devices, return first device
        if selected is None and self.options:
            if self.options[0] != "No devices available":
                return self.options[0]
            else:
                return "No devices available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No devices available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected device."""
        if option in self.options and option != "No devices available":
            self.coordinator.set_selected_device(option)
            # Trigger coordinator update to fetch groups for new device
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected device."""
        current = self.current_option
        if current and current != "No devices available" and self.coordinator.data is not None:
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
        return groups if groups else ["No groups available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected group."""
        selected = self.coordinator.get_selected_group()
        
        # If nothing selected and we have groups, return first group
        if selected is None and self.options:
            if self.options[0] != "No groups available":
                return self.options[0]
            else:
                return "No groups available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No groups available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected group."""
        if option in self.options and option != "No groups available":
            self.coordinator.set_selected_group(option)
            # Trigger coordinator update to fetch members for new group
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected group."""
        current = self.current_option
        if current and current != "No groups available" and self.coordinator.data is not None:
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
        return users if users else ["No users available"]

    @property
    def current_option(self) -> str | None:
        """Return the currently selected user."""
        selected = self.coordinator.get_selected_user()
        
        # If nothing selected and we have users, return first user
        if selected is None and self.options:
            if self.options[0] != "No users available":
                return self.options[0]
            else:
                return "No users available"
        
        # Ensure selected is in options, otherwise return first option
        if selected not in self.options:
            return self.options[0] if self.options else "No users available"
        
        return selected

    async def async_select_option(self, option: str) -> None:
        """Change the selected user."""
        if option in self.options and option != "No users available":
            self.coordinator.set_selected_user(option)
            # Trigger coordinator update to fetch devices for new user
            await self.coordinator.async_request_refresh()
            self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes for the selected user."""
        current = self.current_option
        if current and current != "No users available" and self.coordinator.data is not None:
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
