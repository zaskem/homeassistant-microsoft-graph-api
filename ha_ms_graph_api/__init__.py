"""The Microsoft Graph API Sandbox integration."""

from __future__ import annotations

import logging

import voluptuous as vol

from homeassistant.config_entries import ConfigEntry
from homeassistant.const import Platform
from homeassistant.core import HomeAssistant, ServiceCall
from homeassistant.helpers.aiohttp_client import async_get_clientsession
import homeassistant.helpers.config_validation as cv

from .api import msGraphApiClient
from .const import (
    CONF_CLIENT_ID,
    CONF_CLIENT_SECRET,
    CONF_CLIENT_CERT_PATH,
    CONF_TENANT_ID,
    CONF_UPDATE_INTERVAL,
    CONF_PRIVACY_MODE,
    CONF_SAFE_MODE,
    CONF_USE_CERT_AUTH,
    DEFAULT_UPDATE_INTERVAL,
    DEFAULT_PRIVACY_MODE,
    DEFAULT_SAFE_MODE,
    DEFAULT_USE_CERT_AUTH,
    DOMAIN,
)

_LOGGER = logging.getLogger(__name__)

_PLATFORMS: list[Platform] = [Platform.BINARY_SENSOR, Platform.SENSOR, Platform.SELECT, Platform.TEXT]

type GraphAPIConfigEntry = ConfigEntry[msGraphApiClient]

# Service constants
SERVICE_UPDATE_EXTENSION_ATTR = "update_device_extension_attribute"
SERVICE_UPDATE_USER_PROPERTIES = "update_user_properties"
ATTR_DEVICE_NAME = "device_name"
ATTR_ATTRIBUTE_NUMBER = "attribute_number"
ATTR_VALUE = "value"
ATTR_USER_NAME = "user_name"
ATTR_EMPLOYEE_ID = "employee_id"
ATTR_JOB_TITLE = "job_title"
ATTR_DEPARTMENT = "department"

SERVICE_UPDATE_EXTENSION_ATTR_SCHEMA = vol.Schema(
    {
        vol.Required(ATTR_DEVICE_NAME): cv.string,
        vol.Required(ATTR_ATTRIBUTE_NUMBER): vol.All(
            vol.Coerce(int), vol.Range(min=1, max=15)
        ),
        vol.Optional(ATTR_VALUE): vol.Any(cv.string, None),
    }
)

SERVICE_UPDATE_USER_PROPERTIES_SCHEMA = vol.Schema(
    {
        vol.Required(ATTR_USER_NAME): cv.string,
        vol.Optional(ATTR_EMPLOYEE_ID): vol.Any(cv.string, None),
        vol.Optional(ATTR_JOB_TITLE): vol.Any(cv.string, None),
        vol.Optional(ATTR_DEPARTMENT): vol.Any(cv.string, None),
    }
)


async def async_setup_entry(hass: HomeAssistant, entry: GraphAPIConfigEntry) -> bool:
    """Set up Microsoft Graph API Sandbox from a config entry."""

    session = async_get_clientsession(hass)
    
    # Use options if available, otherwise fall back to data
    update_interval = entry.options.get(
        CONF_UPDATE_INTERVAL,
        entry.data.get(CONF_UPDATE_INTERVAL, DEFAULT_UPDATE_INTERVAL)
    )
    
    # Determine authentication method
    use_cert_auth = entry.options.get(
        CONF_USE_CERT_AUTH,
        entry.data.get(CONF_USE_CERT_AUTH, DEFAULT_USE_CERT_AUTH)
    )
    
    # Get credentials based on authentication method
    if use_cert_auth:
        client_cert_path = entry.options.get(
            CONF_CLIENT_CERT_PATH,
            entry.data.get(CONF_CLIENT_CERT_PATH)
        )
        if not client_cert_path:
            _LOGGER.error(
                "Certificate path not configured. Please reconfigure the integration "
                "with a valid certificate path when using certificate authentication."
            )
            return False
        client_secret = None
    else:
        client_secret = entry.options.get(
            CONF_CLIENT_SECRET,
            entry.data.get(CONF_CLIENT_SECRET)
        )
        if not client_secret:
            _LOGGER.error(
                "Client secret not configured. Please reconfigure the integration "
                "with a valid client secret when using client secret authentication."
            )
            return False
        client_cert_path = None
    
    client = msGraphApiClient(
        client_id=entry.data[CONF_CLIENT_ID],
        tenant_id=entry.data[CONF_TENANT_ID],
        session=session,
        update_interval=update_interval,
        client_secret=client_secret,
        client_cert_path=client_cert_path,
        use_cert_auth=use_cert_auth,
    )
    
    # Authenticate with Microsoft Graph API
    if not await client.authenticate():
        return False
    
    # Store the API client for platforms to access
    entry.runtime_data = client
    
    # Store privacy mode setting in hass.data for sensor platform access
    privacy_mode = entry.options.get(
        CONF_PRIVACY_MODE,
        entry.data.get(CONF_PRIVACY_MODE, DEFAULT_PRIVACY_MODE)
    )
    hass.data.setdefault(DOMAIN, {})
    hass.data[DOMAIN][f"{entry.entry_id}_privacy_mode"] = privacy_mode

    await hass.config_entries.async_forward_entry_setups(entry, _PLATFORMS)
    
    # Register an update listener to reload when options change
    entry.async_on_unload(entry.add_update_listener(async_reload_entry))
    
    # Register service to update extension attributes (only if safe mode is disabled)
    safe_mode = entry.options.get(
        CONF_SAFE_MODE,
        entry.data.get(CONF_SAFE_MODE, DEFAULT_SAFE_MODE)
    )
    
    if not safe_mode:
        async def handle_update_extension_attr(call: ServiceCall) -> None:
            """Handle the service call to update a device extension attribute."""
            device_name = call.data[ATTR_DEVICE_NAME]
            attribute_number = call.data[ATTR_ATTRIBUTE_NUMBER]
            value = call.data.get(ATTR_VALUE)
            
            # Get the device coordinator to find the device object ID
            device_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_coordinator")
            if not device_coordinator or not device_coordinator.data:
                _LOGGER.error("Device coordinator not available")
                return
            
            device_dict = device_coordinator.data.get("device_dict", {})
            if device_name not in device_dict:
                _LOGGER.error("Device '%s' not found", device_name)
                return
            
            device_object_id = device_dict[device_name].get("id")
            if not device_object_id:
                _LOGGER.error("Could not find object ID for device '%s'", device_name)
                return
            
            # Call the API to update the extension attribute
            success = await client.update_device_extension_attribute(
                device_object_id, attribute_number, value
            )
            
            if success:
                # Trigger a coordinator refresh to update the sensor
                await device_coordinator.async_request_refresh()
                _LOGGER.info(
                    "Updated extensionAttribute%d on device '%s' to '%s'",
                    attribute_number,
                    device_name,
                    value,
                )
            else:
                _LOGGER.error(
                    "Failed to update extensionAttribute%d on device '%s'",
                    attribute_number,
                    device_name,
                )
        
        hass.services.async_register(
            DOMAIN,
            SERVICE_UPDATE_EXTENSION_ATTR,
            handle_update_extension_attr,
            schema=SERVICE_UPDATE_EXTENSION_ATTR_SCHEMA,
        )
        
        async def handle_update_user_properties(call: ServiceCall) -> None:
            """Handle the service call to update user properties."""
            user_name = call.data[ATTR_USER_NAME]
            employee_id = call.data.get(ATTR_EMPLOYEE_ID)
            job_title = call.data.get(ATTR_JOB_TITLE)
            department = call.data.get(ATTR_DEPARTMENT)
            
            # Ensure at least one property is provided
            if employee_id is None and job_title is None and department is None:
                _LOGGER.error("At least one property must be provided to update")
                return
            
            # Get the user coordinator to find the user ID
            user_coordinator = hass.data.get(f"{DOMAIN}_{entry.entry_id}_user_coordinator")
            if not user_coordinator or not user_coordinator.data:
                _LOGGER.error("User coordinator not available")
                return
            
            user_dict = user_coordinator.data.get("user_dict", {})
            if user_name not in user_dict:
                _LOGGER.error("User '%s' not found", user_name)
                return
            
            user_id = user_dict[user_name].get("id")
            if not user_id:
                _LOGGER.error("Could not find ID for user '%s'", user_name)
                return
            
            # Call the API to update the user properties
            success = await client.update_user_properties(
                user_id, employee_id, job_title, department
            )
            
            if success:
                # Trigger a coordinator refresh to update the sensors
                await user_coordinator.async_request_refresh()
                updated_props = []
                if employee_id is not None:
                    updated_props.append(f"employeeId={employee_id}")
                if job_title is not None:
                    updated_props.append(f"jobTitle={job_title}")
                if department is not None:
                    updated_props.append(f"department={department}")
                _LOGGER.info(
                    "Updated user '%s': %s",
                    user_name,
                    ", ".join(updated_props),
                )
            else:
                _LOGGER.error(
                    "Failed to update properties for user '%s'",
                    user_name,
                )
        
        hass.services.async_register(
            DOMAIN,
            SERVICE_UPDATE_USER_PROPERTIES,
            handle_update_user_properties,
            schema=SERVICE_UPDATE_USER_PROPERTIES_SCHEMA,
        )

    return True


async def async_reload_entry(hass: HomeAssistant, entry: GraphAPIConfigEntry) -> None:
    """Reload config entry when options change."""
    await hass.config_entries.async_reload(entry.entry_id)


async def async_unload_entry(hass: HomeAssistant, entry: GraphAPIConfigEntry) -> bool:
    """Unload a config entry."""
    # Unregister services if they were registered
    if hass.services.has_service(DOMAIN, SERVICE_UPDATE_EXTENSION_ATTR):
        hass.services.async_remove(DOMAIN, SERVICE_UPDATE_EXTENSION_ATTR)
    if hass.services.has_service(DOMAIN, SERVICE_UPDATE_USER_PROPERTIES):
        hass.services.async_remove(DOMAIN, SERVICE_UPDATE_USER_PROPERTIES)
    
    # Clean up coordinator from hass.data
    hass.data.pop(f"{DOMAIN}_{entry.entry_id}_coordinator", None)
    hass.data.pop(f"{DOMAIN}_{entry.entry_id}_group_coordinator", None)
    hass.data.pop(f"{DOMAIN}_{entry.entry_id}_user_coordinator", None)
    
    # Clean up privacy mode setting
    if DOMAIN in hass.data:
        hass.data[DOMAIN].pop(f"{entry.entry_id}_privacy_mode", None)
    
    return await hass.config_entries.async_unload_platforms(entry, _PLATFORMS)
