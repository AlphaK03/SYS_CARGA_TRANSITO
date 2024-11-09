from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.http.request_options import RequestOptions
import logging

logger = logging.getLogger(__name__)

class TokenAuthenticationContext(AuthenticationContext):
    def __init__(self, site_url, access_token):
        super().__init__(site_url)
        self.access_token = access_token

    def authenticate_request(self, request):
        """
        Adds the access token to the Authorization header for the request.
        """
        request.ensure_header("Authorization", f"Bearer {self.access_token}")
        logger.debug("Authorization header added with access token.")
