import os
from msal import ConfidentialClientApplication
from dotenv import load_dotenv


"""
Generates the access tokens necessary for querying the Microsoft API's
"""
def get_access_token(scope):
    load_dotenv()
    TENANT_ID = os.getenv("TENANT_ID")
    CLIENT_ID = os.getenv("CLIENT_ID")
    CLIENT_SECRET = os.getenv("CLIENT_SECRET")
    authority = f'https://login.microsoftonline.com/{TENANT_ID}'
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )
    result = app.acquire_token_for_client(scopes=scope)
    if 'access_token' in result:
        return result['access_token']
    else:
        raise Exception(f"Failed to get token: {result.get('error_description')}")
    