"""Binary sensor platform for Microsoft Graph API integration."""

from __future__ import annotations

import logging

from homeassistant.components.binary_sensor import BinarySensorEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback

from .const import DOMAIN

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up Microsoft Graph API binary sensor platform."""
    # Get configuration settings
    privacy_mode = hass.data.get(DOMAIN, {}).get(f"{entry.entry_id}_privacy_mode", True)
    safe_mode = entry.options.get(
        "safe_mode",
        entry.data.get("safe_mode", True)
    )
    
    # Add binary sensor entities
    binary_sensors = [
        GraphAPIPrivacyModeSensor(entry, privacy_mode),
        GraphAPISafeModeSensor(entry, safe_mode),
    ]
    
    async_add_entities(binary_sensors, True)


class GraphAPIPrivacyModeSensor(BinarySensorEntity):
    """Binary sensor showing privacy mode configuration status.
    
    This read-only sensor reflects the current privacy mode setting.
    When ON, sensitive data (BitLocker keys, user details) is hidden.
    """

    def __init__(self, entry: ConfigEntry, privacy_mode: bool) -> None:
        """Initialize the binary sensor."""
        self._entry = entry
        self._attr_name = "Graph API Privacy Mode"
        self._attr_unique_id = f"{DOMAIN}_{entry.entry_id}_privacy_mode_status"
        self._attr_icon = "mdi:shield-lock"
        self._attr_is_on = privacy_mode
        self._attr_device_class = None

    @property
    def extra_state_attributes(self):
        """Return additional attributes."""
        return {
            "description": "Privacy Mode hides sensitive data like BitLocker keys and user personal information",
            "configurable": True,
            "config_location": "Integration configuration options",
        }


class GraphAPISafeModeSensor(BinarySensorEntity):
    """Binary sensor showing safe mode configuration status.
    
    This read-only sensor reflects the current safe mode setting.
    When ON, write operations to Entra ID/Intune are disabled.
    """

    def __init__(self, entry: ConfigEntry, safe_mode: bool) -> None:
        """Initialize the binary sensor."""
        self._entry = entry
        self._attr_name = "Graph API Safe Mode"
        self._attr_unique_id = f"{DOMAIN}_{entry.entry_id}_safe_mode_status"
        self._attr_icon = "mdi:shield-check"
        self._attr_is_on = safe_mode
        self._attr_device_class = None

    @property
    def extra_state_attributes(self):
        """Return additional attributes."""
        return {
            "description": "Safe Mode prevents write operations to Entra ID/Intune",
            "configurable": True,
            "config_location": "Integration configuration options",
        }
