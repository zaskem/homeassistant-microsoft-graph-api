"""Microsoft Graph API client."""

from __future__ import annotations

import asyncio
import base64
import hashlib
import logging
import os
import time
import uuid
from typing import Any

import aiohttp
import jwt
from aiohttp import ClientError, ClientSession
from cryptography import x509
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.backends import default_backend

from .const import CONF_CLIENT_ID, CONF_CLIENT_SECRET, CONF_TENANT_ID, CONF_UPDATE_INTERVAL, GRAPH_DEVICES_URL, GRAPH_TOKEN_URL

_LOGGER = logging.getLogger(__name__)


class msGraphApiClient:
    """Client to interact with Microsoft Graph API."""

    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        session: ClientSession,
        update_interval: int,
        client_secret: str | None = None,
        client_cert_path: str | None = None,
        use_cert_auth: bool = False,
    ) -> None:
        """Initialize the API client.
        
        Args:
            client_id: Azure AD application (client) ID
            tenant_id: Azure AD tenant ID
            session: aiohttp ClientSession
            update_interval: Update interval in seconds
            client_secret: Azure AD client secret (for secret-based auth)
            client_cert_path: Path to certificate file in PEM format (for cert-based auth)
            use_cert_auth: Whether to use certificate authentication instead of client secret
        """
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.session = session
        self.update_interval = update_interval
        self.use_cert_auth = use_cert_auth
        
        # Store authentication credentials based on method
        if use_cert_auth:
            if not client_cert_path:
                raise ValueError("client_cert_path is required when use_cert_auth is True")
            self.client_cert_path = client_cert_path
            self.client_secret = None
            self._private_key = None
            self._certificate = None
            self._cert_thumbprint = None
        else:
            if not client_secret:
                raise ValueError("client_secret is required when use_cert_auth is False")
            self.client_secret = client_secret
            self.client_cert_path = None
        
        self._bearer_token: str | None = None

    async def _load_certificate_and_key(self) -> tuple[bytes, bytes]:
        """Load and parse the certificate and private key from file.
        
        Returns:
            Tuple of (private_key_bytes, certificate_der_bytes)
        """
        if self._private_key is None:
            try:
                # Check if file exists first
                if not os.path.exists(self.client_cert_path):
                    _LOGGER.error(
                        "Certificate file not found at '%s'. "
                        "Please ensure the file exists and the path is correct. "
                        "In Home Assistant, certificates are typically stored in /config/ssl/ or /ssl/",
                        self.client_cert_path
                    )
                    raise FileNotFoundError(f"Certificate file not found: {self.client_cert_path}")
                
                # Read file in executor to avoid blocking event loop
                def read_cert_file():
                    with open(self.client_cert_path, "rb") as cert_file:
                        return cert_file.read()
                
                cert_data = await asyncio.get_event_loop().run_in_executor(None, read_cert_file)
                    
                # Try to load as PEM format private key
                try:
                    private_key = serialization.load_pem_private_key(
                        cert_data,
                        password=None,
                        backend=default_backend()
                    )
                    self._private_key = private_key.private_bytes(
                        encoding=serialization.Encoding.PEM,
                        format=serialization.PrivateFormat.PKCS8,
                        encryption_algorithm=serialization.NoEncryption()
                    )
                    
                    # Also try to load the certificate from the same file
                    try:
                        certificate = x509.load_pem_x509_certificate(
                            cert_data,
                            backend=default_backend()
                        )
                        self._certificate = certificate.public_bytes(serialization.Encoding.DER)
                        
                        # Calculate SHA-1 thumbprint for x5t header
                        thumbprint = hashlib.sha1(self._certificate).digest()
                        self._cert_thumbprint = base64.urlsafe_b64encode(thumbprint).decode('utf-8').rstrip('=')
                        
                        _LOGGER.debug("Successfully loaded certificate with thumbprint: %s", self._cert_thumbprint)
                    except Exception as cert_err:
                        _LOGGER.warning(
                            "Could not load certificate from %s (only found private key): %s. "
                            "For certificate authentication, the file should contain both the private key and certificate.",
                            self.client_cert_path, cert_err
                        )
                        raise ValueError("Certificate file must contain both private key and certificate in PEM format")
                        
                except Exception as parse_err:
                    _LOGGER.error("Failed to parse certificate/key file: %s", parse_err)
                    raise
                    
            except FileNotFoundError:
                raise
            except Exception as err:
                _LOGGER.error("Failed to load certificate: %s", err)
                raise
                
        return self._private_key, self._certificate

    async def _create_client_assertion(self) -> str:
        """Create a JWT client assertion for certificate-based authentication.
        
        Returns:
            A signed JWT assertion string
        """
        # Load private key and certificate
        private_key, _ = await self._load_certificate_and_key()
        
        # JWT header and claims
        now = int(time.time())
        token_url = GRAPH_TOKEN_URL.format(self.tenant_id)
        
        payload = {
            "aud": token_url,
            "exp": now + 600,  # Token expires in 10 minutes
            "iss": self.client_id,
            "jti": str(uuid.uuid4()),  # Unique identifier for this JWT
            "nbf": now,
            "sub": self.client_id,
            "iat": now,
        }
        
        # Create headers with certificate thumbprint (required by Azure AD)
        headers = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": self._cert_thumbprint  # Certificate SHA-1 thumbprint
        }
        
        # Sign the JWT with RS256 algorithm (RSA with SHA-256)
        assertion = jwt.encode(
            payload,
            private_key,
            algorithm="RS256",
            headers=headers,
        )
        
        return assertion

    async def authenticate(self) -> bool:
        """Authenticate with Microsoft Graph API and obtain a bearer token."""
        try:
            token_url = GRAPH_TOKEN_URL.format(self.tenant_id)
            
            if self.use_cert_auth:
                # Certificate-based authentication
                client_assertion = await self._create_client_assertion()
                
                data = {
                    "grant_type": "client_credentials",
                    "client_id": self.client_id,
                    "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                    "client_assertion": client_assertion,
                    "scope": "https://graph.microsoft.com/.default",
                }
            else:
                # Client secret authentication
                data = {
                    "grant_type": "client_credentials",
                    "client_id": self.client_id,
                    "client_secret": self.client_secret,
                    "scope": "https://graph.microsoft.com/.default",
                }
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded",
            }
            
            async with self.session.post(
                token_url, data=data, headers=headers
            ) as response:
                if response.status != 200:
                    error_text = await response.text()
                    _LOGGER.error(
                        "Failed to authenticate with Microsoft Graph API: %s - %s",
                        response.status,
                        error_text,
                    )
                    return False
                
                result = await response.json()
                self._bearer_token = result.get("access_token")
                
                if not self._bearer_token:
                    _LOGGER.error("No access token received from Microsoft Graph API")
                    return False
                
                auth_method = "certificate" if self.use_cert_auth else "client secret"
                _LOGGER.debug("Successfully authenticated with Microsoft Graph API using %s", auth_method)
                return True
                
        except Exception as err:
            _LOGGER.error("Error authenticating with Microsoft Graph API: %s", err)
            return False

    async def get_devices(self) -> list[dict[str, Any]]:
        """Fetch devices from Microsoft Graph API."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch devices: authentication failed")
                return []
        
        try:
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(
                GRAPH_DEVICES_URL, headers=headers
            ) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        # Retry with new token
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(
                            GRAPH_DEVICES_URL, headers=headers
                        ) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                devices = result.get("value", [])
                                _LOGGER.debug("Fetched %d devices", len(devices))
                                return devices
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch devices from Microsoft Graph API: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                devices = result.get("value", [])
                _LOGGER.debug("Fetched %d devices", len(devices))
                return devices
                
        except ClientError as err:
            _LOGGER.error("Error fetching devices from Microsoft Graph API: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching devices: %s", err)
            return []

    async def get_device_groups(self, device_object_id: str) -> list[str]:
        """Fetch groups that a device is a member of."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch device groups: authentication failed")
                return []
        
        try:
            url = f"https://graph.microsoft.com/v1.0/devices/{device_object_id}/memberOf/$/microsoft.graph.group?$select=displayName"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(url, headers=headers) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                groups = [g.get("displayName", "") for g in result.get("value", [])]
                                _LOGGER.debug("Fetched %d groups for device", len(groups))
                                return groups
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch device groups: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                groups = [g.get("displayName", "") for g in result.get("value", [])]
                _LOGGER.debug("Fetched %d groups for device", len(groups))
                return groups
                
        except ClientError as err:
            _LOGGER.error("Error fetching device groups: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching device groups: %s", err)
            return []

    async def get_bitlocker_recovery_keys(self, device_id: str) -> list[str]:
        """Fetch BitLocker recovery keys for a device (two-step process).
        
        Step 1: Get list of recovery key IDs for the device
        Step 2: Fetch each recovery key value
        """
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch BitLocker keys: authentication failed")
                return ["Authentication failed"]
        
        try:
            # Step 1: Get recovery key IDs for the device
            url = f"https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys?$filter=deviceId eq '{device_id}'"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if not await self.authenticate():
                        return ["Token expired"]
                    headers["Authorization"] = f"Bearer {self._bearer_token}"
                    # Retry after re-authentication
                    async with self.session.get(url, headers=headers) as retry_response:
                        if retry_response.status != 200:
                            return ["Failed to fetch recovery key list"]
                        result = await retry_response.json()
                else:
                    if response.status != 200:
                        _LOGGER.error(
                            "Failed to fetch BitLocker recovery keys list: %s",
                            response.status,
                        )
                        return ["Failed to fetch recovery key list"]
                    result = await response.json()
            
            # Check if we got any recovery keys
            recovery_key_ids = result.get("value", [])
            if not recovery_key_ids:
                return ["No keys available"]
            
            # Step 2: Fetch each recovery key value
            recovery_keys = []
            for key_info in recovery_key_ids:
                key_id = key_info.get("id")
                if not key_id:
                    continue
                
                key_url = f"https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/{key_id}?$select=key"
                
                async with self.session.get(key_url, headers=headers) as key_response:
                    if key_response.status == 200:
                        key_result = await key_response.json()
                        key_value = key_result.get("key")
                        if key_value:
                            recovery_keys.append(key_value)
                    else:
                        _LOGGER.warning(
                            "Failed to fetch recovery key %s: %s",
                            key_id,
                            key_response.status,
                        )
            
            if not recovery_keys:
                return ["Keys found but could not retrieve values"]
            
            _LOGGER.debug("Fetched %d BitLocker recovery keys", len(recovery_keys))
            return recovery_keys
                
        except ClientError as err:
            _LOGGER.error("Error fetching BitLocker recovery keys: %s", err)
            return ["Network error"]
        except Exception as err:
            _LOGGER.error("Unexpected error fetching BitLocker keys: %s", err)
            return ["Unexpected error"]

    async def get_groups(self) -> list[dict[str, Any]]:
        """Fetch security-enabled groups from Microsoft Graph API."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch groups: authentication failed")
                return []
        
        try:
            # Fetch only security-enabled groups with specific fields
            url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName,securityEnabled,groupTypes,createdDateTime&$filter=(securityEnabled eq true)"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(url, headers=headers) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                groups = result.get("value", [])
                                _LOGGER.debug("Fetched %d groups", len(groups))
                                return groups
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch groups: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                groups = result.get("value", [])
                _LOGGER.debug("Fetched %d groups", len(groups))
                return groups
                
        except ClientError as err:
            _LOGGER.error("Error fetching groups: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching groups: %s", err)
            return []

    async def get_group_members(self, group_id: str) -> list[dict[str, Any]]:
        """Fetch transitive members (devices) of a group."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch group members: authentication failed")
                return []
        
        try:
            url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/transitiveMembers?$select=id,displayName,deviceId"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
                "Content-Type": "application/json",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(url, headers=headers) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                # Filter to only device members (those with deviceId)
                                all_members = result.get("value", [])
                                device_members = [m for m in all_members if m.get("deviceId")]
                                _LOGGER.debug("Fetched %d device members for group", len(device_members))
                                return device_members
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch group members: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                # Filter to only device members (those with deviceId)
                all_members = result.get("value", [])
                device_members = [m for m in all_members if m.get("deviceId")]
                _LOGGER.debug("Fetched %d device members for group", len(device_members))
                return device_members
                
        except ClientError as err:
            _LOGGER.error("Error fetching group members: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching group members: %s", err)
            return []

    async def get_users(self) -> list[dict[str, Any]]:
        """Fetch users from Microsoft Graph API."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch users: authentication failed")
                return []
        
        try:
            # Fetch users with specific fields (id, displayName, mail, userPrincipalName)
            url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,employeeId,jobTitle,department"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(url, headers=headers) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                users = result.get("value", [])
                                _LOGGER.debug("Fetched %d users", len(users))
                                return users
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch users: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                users = result.get("value", [])
                _LOGGER.debug("Fetched %d users", len(users))
                return users
                
        except ClientError as err:
            _LOGGER.error("Error fetching users: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching users: %s", err)
            return []

    async def get_user_devices(self, user_id: str) -> list[str]:
        """Fetch devices owned by a user."""
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot fetch user devices: authentication failed")
                return []
        
        try:
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/ownedDevices?$select=displayName"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
            }
            
            async with self.session.get(url, headers=headers) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.get(url, headers=headers) as retry_response:
                            if retry_response.status == 200:
                                result = await retry_response.json()
                                devices = [d.get("displayName", "") for d in result.get("value", [])]
                                _LOGGER.debug("Fetched %d devices for user", len(devices))
                                return devices
                    return []
                
                if response.status != 200:
                    _LOGGER.error(
                        "Failed to fetch user devices: %s",
                        response.status,
                    )
                    return []
                
                result = await response.json()
                devices = [d.get("displayName", "") for d in result.get("value", [])]
                _LOGGER.debug("Fetched %d devices for user", len(devices))
                return devices
                
        except ClientError as err:
            _LOGGER.error("Error fetching user devices: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Unexpected error fetching user devices: %s", err)
            return []

    async def test_connection(self) -> bool:
        """Test the connection to Microsoft Graph API."""
        return await self.authenticate()

    async def update_device_extension_attribute(
        self, device_object_id: str, attribute_number: int, value: str | None
    ) -> bool:
        """Update a single extension attribute on a device.
        
        Args:
            device_object_id: The Azure AD object ID of the device
            attribute_number: The extension attribute number (1-15)
            value: The value to set (None to clear)
            
        Returns:
            True if successful, False otherwise
        """
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot update extension attribute: authentication failed")
                return False
        
        if not 1 <= attribute_number <= 15:
            _LOGGER.error("Extension attribute number must be between 1 and 15")
            return False
        
        try:
            url = f"https://graph.microsoft.com/v1.0/devices/{device_object_id}"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
                "Content-Type": "application/json",
            }
            
            payload = {
                "extensionAttributes": {
                    f"extensionAttribute{attribute_number}": value
                }
            }
            
            async with self.session.patch(
                url, json=payload, headers=headers
            ) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.patch(
                            url, json=payload, headers=headers
                        ) as retry_response:
                            if retry_response.status == 204:
                                _LOGGER.info(
                                    "Successfully updated extensionAttribute%d on device %s",
                                    attribute_number,
                                    device_object_id,
                                )
                                return True
                            else:
                                _LOGGER.error(
                                    "Failed to update extension attribute: %s",
                                    retry_response.status,
                                )
                                return False
                    return False
                
                if response.status == 204:
                    _LOGGER.info(
                        "Successfully updated extensionAttribute%d on device %s",
                        attribute_number,
                        device_object_id,
                    )
                    return True
                else:
                    error_text = await response.text()
                    _LOGGER.error(
                        "Failed to update extension attribute: %s - %s",
                        response.status,
                        error_text,
                    )
                    return False
                
        except ClientError as err:
            _LOGGER.error("Error updating extension attribute: %s", err)
            return False
        except Exception as err:
            _LOGGER.error("Unexpected error updating extension attribute: %s", err)
            return False

    async def update_user_properties(
        self, user_id: str, employee_id: str | None = None, job_title: str | None = None, department: str | None = None
    ) -> bool:
        """Update user properties (employee ID, job title, and/or department).
        
        Args:
            user_id: The user ID or userPrincipalName
            employee_id: The employee ID to set (None to leave unchanged)
            job_title: The job title to set (None to leave unchanged)
            department: The department to set (None to leave unchanged)
            
        Returns:
            True if successful, False otherwise
        """
        if not self._bearer_token:
            if not await self.authenticate():
                _LOGGER.error("Cannot update user properties: authentication failed")
                return False
        
        # Build payload with only the properties that were provided
        payload = {}
        if employee_id is not None:
            payload["employeeId"] = employee_id
        if job_title is not None:
            payload["jobTitle"] = job_title
        if department is not None:
            payload["department"] = department
        
        if not payload:
            _LOGGER.error("No user properties provided to update")
            return False
        
        try:
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
            headers = {
                "Authorization": f"Bearer {self._bearer_token}",
                "Content-Type": "application/json",
            }
            
            async with self.session.patch(
                url, json=payload, headers=headers
            ) as response:
                if response.status == 401:
                    # Token expired, re-authenticate
                    _LOGGER.debug("Token expired, re-authenticating")
                    if await self.authenticate():
                        headers["Authorization"] = f"Bearer {self._bearer_token}"
                        async with self.session.patch(
                            url, json=payload, headers=headers
                        ) as retry_response:
                            if retry_response.status == 204:
                                _LOGGER.info(
                                    "Successfully updated user properties for %s",
                                    user_id,
                                )
                                return True
                            else:
                                _LOGGER.error(
                                    "Failed to update user properties: %s",
                                    retry_response.status,
                                )
                                return False
                    return False
                
                if response.status == 204:
                    _LOGGER.info(
                        "Successfully updated user properties for %s",
                        user_id,
                    )
                    return True
                else:
                    error_text = await response.text()
                    _LOGGER.error(
                        "Failed to update user properties: %s - %s",
                        response.status,
                        error_text,
                    )
                    return False
                
        except ClientError as err:
            _LOGGER.error("Error updating user properties: %s", err)
            return False
        except Exception as err:
            _LOGGER.error("Unexpected error updating user properties: %s", err)
            return False
