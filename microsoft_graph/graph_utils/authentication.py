import keyring
import logging
import msal

logger = logging.getLogger(__name__)


class Application:
    def __init__(self, tenant_id, application_id, application_secret):
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.client_id = application_id
        b = keyring.get_credential(application_secret, None)
        self.application_secret = b.password
        self.scope = ["https://graph.microsoft.com/.default"]
        self.endpoint = "https://graph.microsoft.com/v1.0/users"
# This is an aplication instance with all the required parameters to connect with graph at a daemon application level

    # This function acquires the connection token for an aplication daemon
    def acquire_token(self):
        
        app = msal.ConfidentialClientApplication(
            self.client_id, authority=self.authority,
            client_credential= self.application_secret,)
        result = None

        result = app.acquire_token_silent(self.scope, account=None)

        if not result:
            logger.info("No suitable token exists in cache. Let's get a new one from AAD.")
            result = app.acquire_token_for_client(scopes=self.scope)
            
        return result
