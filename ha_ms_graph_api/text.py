"""Text platform for Microsoft Graph API integration."""

from __future__ import annotations

import logging
from typing import Any

from homeassistant.components.text import TextEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, callback
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.event import async_track_state_change_event

from .const import DOMAIN

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up Microsoft Graph API text platform."""
    # Get safe mode setting
    safe_mode = entry.options.get(
        "safe_mode",
        entry.data.get("safe_mode", True)
    )

    # Only add text entities when safe mode is disabled
    if not safe_mode:
        entities = []
        
        # Get the device coordinator for extension attributes
        device_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_coordinator")
        if device_coordinator:
            entities.append(GraphAPIExtensionAttributeText(hass, device_coordinator, entry))
        
        # Get the user coordinator for user properties
        user_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_user_coordinator")
        if user_coordinator:
            entities.extend([
                GraphAPIUserEmployeeIDText(hass, user_coordinator, entry),
                GraphAPIUserJobTitleText(hass, user_coordinator, entry),
                GraphAPIUserDepartmentText(hass, user_coordinator, entry),
            ])
        
        if entities:
            async_add_entities(entities, True)


class GraphAPIExtensionAttributeText(TextEntity):
    """Text entity for editing extension attribute values.
    
    This entity displays the value of the currently selected extension attribute
    and allows free-form text editing. Changes are stored locally and can be
    manually saved using the update_device_extension_attribute service.
    """

    def __init__(self, hass: HomeAssistant, coordinator, entry: ConfigEntry) -> None:
        """Initialize the text entity."""
        self._hass = hass
        self._coordinator = coordinator
        self._entry = entry
        self._attr_name = "Graph API Extension Attribute Editor"
        self._attr_unique_id = f"{DOMAIN}_extension_attribute_editor"
        self._attr_icon = "mdi:form-textbox"
        self._attr_native_value = ""
        self._attr_native_max = 1024  # Maximum length for extension attributes
        self._attr_mode = "text"  # Single-line text input
        
        # Track the selector entities
        self._device_selector_entity_id = f"select.{DOMAIN}_device_select"
        self._attr_selector_entity_id = f"select.{DOMAIN}_{entry.entry_id}_extension_attr_select"
        self._unsubscribe = None

    async def async_added_to_hass(self) -> None:
        """Subscribe to state changes when entity is added."""
        await super().async_added_to_hass()
        
        # Subscribe to changes in both selector entities
        self._unsubscribe = async_track_state_change_event(
            self._hass,
            [self._device_selector_entity_id, self._attr_selector_entity_id],
            self._handle_selector_change,
        )
        
        # Set initial value
        await self._update_from_selectors()

    async def async_will_remove_from_hass(self) -> None:
        """Unsubscribe when entity is removed."""
        if self._unsubscribe:
            self._unsubscribe()
        await super().async_will_remove_from_hass()

    @callback
    def _handle_selector_change(self, event) -> None:
        """Handle selector state changes."""
        self._hass.async_create_task(self._update_from_selectors())

    async def _update_from_selectors(self) -> None:
        """Update the text value based on current selector states."""
        # Get the device selector state
        device_selector_state = self._hass.states.get(self._device_selector_entity_id)
        if not device_selector_state:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the extension attribute selector state
        attr_selector_state = self._hass.states.get(self._attr_selector_entity_id)
        if not attr_selector_state:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the selected device
        selected_device = device_selector_state.state
        if not selected_device or selected_device in ["Select Device", "No devices available"]:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the attribute number
        attribute_number = attr_selector_state.attributes.get("attribute_number")
        if not attribute_number:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the value from coordinator data
        if self._coordinator.data is None:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        device_dict = self._coordinator.data.get("device_dict", {})
        if selected_device not in device_dict:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        device_data = device_dict[selected_device]
        extension_attrs = device_data.get("extensionAttributes", {})
        
        # Get the specific extension attribute value
        attr_key = f"extensionAttribute{attribute_number}"
        value = extension_attrs.get(attr_key)
        
        # Set the value (empty string if None)
        self._attr_native_value = str(value) if value is not None else ""
        self.async_write_ha_state()

    async def async_set_value(self, value: str) -> None:
        """Update the text value."""
        self._attr_native_value = value
        self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        # Get selector states
        device_selector_state = self._hass.states.get(self._device_selector_entity_id)
        attr_selector_state = self._hass.states.get(self._attr_selector_entity_id)
        
        attributes = {
            "selected_device": None,
            "attribute_number": None,
            "attribute_label": None,
            "editable": True,
        }
        
        if device_selector_state:
            attributes["selected_device"] = device_selector_state.state
        
        if attr_selector_state:
            attributes["attribute_number"] = attr_selector_state.attributes.get("attribute_number")
            attributes["attribute_label"] = attr_selector_state.state
        
        return attributes


class GraphAPIUserEmployeeIDText(TextEntity):
    """Text entity for editing user employee ID.
    
    This entity displays the employee ID of the currently selected user
    and allows free-form text editing. Changes are stored locally and can be
    manually saved using the update_user_properties service.
    """

    def __init__(self, hass: HomeAssistant, coordinator, entry: ConfigEntry) -> None:
        """Initialize the text entity."""
        self._hass = hass
        self._coordinator = coordinator
        self._entry = entry
        self._attr_name = "Graph API User Employee ID Editor"
        self._attr_unique_id = f"{DOMAIN}_user_employee_id_editor"
        self._attr_icon = "mdi:badge-account-outline"
        self._attr_native_value = ""
        self._attr_native_max = 256
        self._attr_mode = "text"
        
        # Track the user selector entity
        self._user_selector_entity_id = f"select.{DOMAIN}_user_select"
        self._unsubscribe = None

    async def async_added_to_hass(self) -> None:
        """Subscribe to state changes when entity is added."""
        await super().async_added_to_hass()
        
        # Subscribe to changes in user selector
        self._unsubscribe = async_track_state_change_event(
            self._hass,
            [self._user_selector_entity_id],
            self._handle_selector_change,
        )
        
        # Set initial value
        await self._update_from_selector()

    async def async_will_remove_from_hass(self) -> None:
        """Unsubscribe when entity is removed."""
        if self._unsubscribe:
            self._unsubscribe()
        await super().async_will_remove_from_hass()

    @callback
    def _handle_selector_change(self, event) -> None:
        """Handle selector state changes."""
        self._hass.async_create_task(self._update_from_selector())

    async def _update_from_selector(self) -> None:
        """Update the text value based on current selector state."""
        # Get the user selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        if not user_selector_state:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the selected user
        selected_user = user_selector_state.state
        if not selected_user or selected_user in ["Select User", "No users available"]:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the value from coordinator data
        if self._coordinator.data is None:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_dict = self._coordinator.data.get("user_dict", {})
        if selected_user not in user_dict:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_data = user_dict[selected_user]
        value = user_data.get("employeeId")
        
        # Set the value (empty string if None)
        self._attr_native_value = str(value) if value is not None else ""
        self.async_write_ha_state()

    async def async_set_value(self, value: str) -> None:
        """Update the text value."""
        self._attr_native_value = value
        self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        # Get selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        
        attributes = {
            "selected_user": None,
            "editable": True,
        }
        
        if user_selector_state:
            attributes["selected_user"] = user_selector_state.state
        
        return attributes


class GraphAPIUserJobTitleText(TextEntity):
    """Text entity for editing user job title.
    
    This entity displays the job title of the currently selected user
    and allows free-form text editing. Changes are stored locally and can be
    manually saved using the update_user_properties service.
    """

    def __init__(self, hass: HomeAssistant, coordinator, entry: ConfigEntry) -> None:
        """Initialize the text entity."""
        self._hass = hass
        self._coordinator = coordinator
        self._entry = entry
        self._attr_name = "Graph API User Job Title Editor"
        self._attr_unique_id = f"{DOMAIN}_user_job_title_editor"
        self._attr_icon = "mdi:briefcase-outline"
        self._attr_native_value = ""
        self._attr_native_max = 256
        self._attr_mode = "text"
        
        # Track the user selector entity
        self._user_selector_entity_id = f"select.{DOMAIN}_user_select"
        self._unsubscribe = None

    async def async_added_to_hass(self) -> None:
        """Subscribe to state changes when entity is added."""
        await super().async_added_to_hass()
        
        # Subscribe to changes in user selector
        self._unsubscribe = async_track_state_change_event(
            self._hass,
            [self._user_selector_entity_id],
            self._handle_selector_change,
        )
        
        # Set initial value
        await self._update_from_selector()

    async def async_will_remove_from_hass(self) -> None:
        """Unsubscribe when entity is removed."""
        if self._unsubscribe:
            self._unsubscribe()
        await super().async_will_remove_from_hass()

    @callback
    def _handle_selector_change(self, event) -> None:
        """Handle selector state changes."""
        self._hass.async_create_task(self._update_from_selector())

    async def _update_from_selector(self) -> None:
        """Update the text value based on current selector state."""
        # Get the user selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        if not user_selector_state:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the selected user
        selected_user = user_selector_state.state
        if not selected_user or selected_user in ["Select User", "No users available"]:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the value from coordinator data
        if self._coordinator.data is None:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_dict = self._coordinator.data.get("user_dict", {})
        if selected_user not in user_dict:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_data = user_dict[selected_user]
        value = user_data.get("jobTitle")
        
        # Set the value (empty string if None)
        self._attr_native_value = str(value) if value is not None else ""
        self.async_write_ha_state()

    async def async_set_value(self, value: str) -> None:
        """Update the text value."""
        self._attr_native_value = value
        self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        # Get selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        
        attributes = {
            "selected_user": None,
            "editable": True,
        }
        
        if user_selector_state:
            attributes["selected_user"] = user_selector_state.state
        
        return attributes


class GraphAPIUserDepartmentText(TextEntity):
    """Text entity for editing user department.
    
    This entity displays the department of the currently selected user
    and allows free-form text editing. Changes are stored locally and can be
    manually saved using the update_user_properties service.
    """

    def __init__(self, hass: HomeAssistant, coordinator, entry: ConfigEntry) -> None:
        """Initialize the text entity."""
        self._hass = hass
        self._coordinator = coordinator
        self._entry = entry
        self._attr_name = "Graph API User Department Editor"
        self._attr_unique_id = f"{DOMAIN}_user_department_editor"
        self._attr_icon = "mdi:office-building-outline"
        self._attr_native_value = ""
        self._attr_native_max = 256
        self._attr_mode = "text"
        
        # Track the user selector entity
        self._user_selector_entity_id = f"select.{DOMAIN}_user_select"
        self._unsubscribe = None

    async def async_added_to_hass(self) -> None:
        """Subscribe to state changes when entity is added."""
        await super().async_added_to_hass()
        
        # Subscribe to changes in user selector
        self._unsubscribe = async_track_state_change_event(
            self._hass,
            [self._user_selector_entity_id],
            self._handle_selector_change,
        )
        
        # Set initial value
        await self._update_from_selector()

    async def async_will_remove_from_hass(self) -> None:
        """Unsubscribe when entity is removed."""
        if self._unsubscribe:
            self._unsubscribe()
        await super().async_will_remove_from_hass()

    @callback
    def _handle_selector_change(self, event) -> None:
        """Handle selector state changes."""
        self._hass.async_create_task(self._update_from_selector())

    async def _update_from_selector(self) -> None:
        """Update the text value based on current selector state."""
        # Get the user selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        if not user_selector_state:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the selected user
        selected_user = user_selector_state.state
        if not selected_user or selected_user in ["Select User", "No users available"]:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        # Get the value from coordinator data
        if self._coordinator.data is None:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_dict = self._coordinator.data.get("user_dict", {})
        if selected_user not in user_dict:
            self._attr_native_value = ""
            self.async_write_ha_state()
            return
        
        user_data = user_dict[selected_user]
        value = user_data.get("department")
        
        # Set the value (empty string if None)
        self._attr_native_value = str(value) if value is not None else ""
        self.async_write_ha_state()

    async def async_set_value(self, value: str) -> None:
        """Update the text value."""
        self._attr_native_value = value
        self.async_write_ha_state()

    @property
    def extra_state_attributes(self) -> dict[str, Any]:
        """Return additional attributes."""
        # Get selector state
        user_selector_state = self._hass.states.get(self._user_selector_entity_id)
        
        attributes = {
            "selected_user": None,
            "editable": True,
        }
        
        if user_selector_state:
            attributes["selected_user"] = user_selector_state.state
        
        return attributes
