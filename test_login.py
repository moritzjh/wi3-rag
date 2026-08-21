import os
from dotenv import load_dotenv
from msal import PublicClientApplication

load_dotenv()

TENANT_ID=os.getenv("TENANT_ID")
LOGIN_CLIENT_ID=os.getenv("LOGIN_CLIENT_ID")

app = PublicClientApplication(
    LOGIN_CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

result = app.acquire_token_interactive(scopes=["User.Read"],
                                       prompt="login")

if "id_token_claims" in result:
    email = result["id_token_claims"].get("preferred_username")
    name = result["id_token_claims"].get("name")
    print(f"Erfolgreich eingeloggt!")
    print(f"E-Mail: {email}")
    print(f"Name: {name}")
else:
    print("Login fehlgeschlagen")
