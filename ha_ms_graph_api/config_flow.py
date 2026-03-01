"""Config flow for the Microsoft Graph API Sandbox integration."""

from __future__ import annotations

import logging
from typing import Any

import aiohttp
import voluptuous as vol

from homeassistant.config_entries import ConfigFlow, ConfigFlowResult, OptionsFlow
from homeassistant.core import HomeAssistant, callback
from homeassistant.exceptions import HomeAssistantError
from homeassistant.helpers.aiohttp_client import async_get_clientsession

from .api import msGraphApiClient
from .const import (
    CONF_CLIENT_ID,
    CONF_CLIENT_SECRET,
    CONF_TENANT_ID,
    CONF_CLIENT_CERT_PATH,
    CONF_UPDATE_INTERVAL,
    CONF_SAFE_MODE,
    CONF_PRIVACY_MODE,
    CONF_USE_CERT_AUTH,
    DEFAULT_UPDATE_INTERVAL,
    DEFAULT_SAFE_MODE,
    DEFAULT_PRIVACY_MODE,
    DEFAULT_USE_CERT_AUTH,
    DOMAIN,
)

_LOGGER = logging.getLogger(__name__)

STEP_USER_DATA_SCHEMA = vol.Schema(
    {
        vol.Required(CONF_CLIENT_ID): str,
        vol.Required(CONF_TENANT_ID): str,
        vol.Optional(CONF_CLIENT_SECRET): str,
        vol.Optional(CONF_USE_CERT_AUTH, default=DEFAULT_USE_CERT_AUTH): bool,
        vol.Optional(CONF_CLIENT_CERT_PATH): str,
        vol.Optional(CONF_UPDATE_INTERVAL, default=DEFAULT_UPDATE_INTERVAL): int,
        vol.Optional(CONF_SAFE_MODE, default=DEFAULT_SAFE_MODE): bool,
        vol.Optional(CONF_PRIVACY_MODE, default=DEFAULT_PRIVACY_MODE): bool,
    }
)


async def validate_input(hass: HomeAssistant, data: dict[str, Any]) -> dict[str, Any]:
    """Validate the user input allows us to connect.

    Data has the keys from STEP_USER_DATA_SCHEMA with values provided by the user.
    """
    session = async_get_clientsession(hass)
    
    # Determine authentication method
    use_cert_auth = data.get(CONF_USE_CERT_AUTH, DEFAULT_USE_CERT_AUTH)
    
    # Validate that required credentials are provided
    if use_cert_auth:
        if not data.get(CONF_CLIENT_CERT_PATH):
            raise ValueError("Certificate path is required when using certificate authentication")
    else:
        if not data.get(CONF_CLIENT_SECRET):
            raise ValueError("Client secret is required when not using certificate authentication")
    
    client = msGraphApiClient(
        client_id=data[CONF_CLIENT_ID],
        tenant_id=data[CONF_TENANT_ID],
        session=session,
        update_interval=data.get(CONF_UPDATE_INTERVAL, DEFAULT_UPDATE_INTERVAL),
        client_secret=data.get(CONF_CLIENT_SECRET),
        client_cert_path=data.get(CONF_CLIENT_CERT_PATH),
        use_cert_auth=use_cert_auth,
    )

    if not await client.test_connection():
        raise InvalidAuth

    # Return info that you want to store in the config entry.
    return {"title": "Microsoft Graph API Sandbox"}


class ConfigFlow(ConfigFlow, domain=DOMAIN):
    """Handle a config flow for Microsoft Graph API Sandbox."""

    VERSION = 1

    @staticmethod
    @callback
    def async_get_options_flow(config_entry):
        """Get the options flow for this handler."""
        return OptionsFlowHandler()

    async def async_step_user(
        self, user_input: dict[str, Any] | None = None
    ) -> ConfigFlowResult:
        """Handle the initial step."""
        errors: dict[str, str] = {}
        if user_input is not None:
            try:
                info = await validate_input(self.hass, user_input)
            except CannotConnect:
                errors["base"] = "cannot_connect"
            except InvalidAuth:
                errors["base"] = "invalid_auth"
            except Exception:
                _LOGGER.exception("Unexpected exception")
                errors["base"] = "unknown"
            else:
                return self.async_create_entry(title=info["title"], data=user_input)

        return self.async_show_form(
            step_id="user", data_schema=STEP_USER_DATA_SCHEMA, errors=errors
        )


class CannotConnect(HomeAssistantError):
    """Error to indicate we cannot connect."""


class InvalidAuth(HomeAssistantError):
    """Error to indicate there is invalid auth."""


class OptionsFlowHandler(OptionsFlow):
    """Handle options flow for Graph API integration."""

    async def async_step_init(self, user_input: dict[str, Any] | None = None):
        """Manage the options."""
        if user_input is not None:
            return self.async_create_entry(title="", data=user_input)

        return self.async_show_form(
            step_id="init",
            data_schema=vol.Schema(
                {
                    vol.Optional(
                        CONF_CLIENT_SECRET,
                        default=self.config_entry.options.get(
                            CONF_CLIENT_SECRET,
                            self.config_entry.data.get(CONF_CLIENT_SECRET, "")
                        ),
                    ): str,
                    vol.Optional(
                        CONF_USE_CERT_AUTH,
                        default=self.config_entry.options.get(
                            CONF_USE_CERT_AUTH,
                            self.config_entry.data.get(CONF_USE_CERT_AUTH, DEFAULT_USE_CERT_AUTH)
                        ),
                    ): bool,
                    vol.Optional(
                        CONF_CLIENT_CERT_PATH,
                        default=self.config_entry.options.get(
                            CONF_CLIENT_CERT_PATH,
                            self.config_entry.data.get(CONF_CLIENT_CERT_PATH, "")
                        ),
                    ): str,
                    vol.Optional(
                        CONF_UPDATE_INTERVAL,
                        default=self.config_entry.options.get(
                            CONF_UPDATE_INTERVAL, 
                            self.config_entry.data.get(CONF_UPDATE_INTERVAL, DEFAULT_UPDATE_INTERVAL)
                        ),
                    ): int,
                    vol.Optional(
                        CONF_SAFE_MODE,
                        default=self.config_entry.options.get(
                            CONF_SAFE_MODE,
                            self.config_entry.data.get(CONF_SAFE_MODE, DEFAULT_SAFE_MODE)
                        ),
                    ): bool,
                    vol.Optional(
                        CONF_PRIVACY_MODE,
                        default=self.config_entry.options.get(
                            CONF_PRIVACY_MODE,
                            self.config_entry.data.get(CONF_PRIVACY_MODE, DEFAULT_PRIVACY_MODE)
                        ),
                    ): bool,
                }
            ),
        )
