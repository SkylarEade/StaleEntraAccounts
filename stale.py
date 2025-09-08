import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta
from token_gen import get_access_token
from export import export_to_xlsx


def get_accounts(header):
    endpoint = ("https://graph.microsoft.com/v1.0/users"
                "?$select=id,displayName,userPrincipalName,department,officeLocation,signInActivity,accountEnabled,employeeId")
    users = []
    while endpoint:
        response = requests.get(endpoint, headers=header)
        if response.status_code == 200:
            data = response.json()
            users.extend(data.get('value', []))
            endpoint = data.get('@odata.nextLink')
        else:
            raise Exception (f"Error {response.status_code}: {response.text}")
    return users

def stale_accounts(users, headers):
    stale = []
    for user in users:
        if user.get("accountEnabled", False):
            last_seen = get_last_seen(user)
            if last_seen:
                user_date = last_seen
                if user_date < datetime.today() - relativedelta(days=90):
                    user["hasLicense"] = has_license(user["id"], headers)
                    stale.append(user)
    return stale

def get_last_seen(user):
    logs = user.get("signInActivity")
    if logs:
        interactive = logs.get("lastSuccessfulSignInDateTime")
        non_interactive = logs.get("lastNonInteractiveSignInDateTime")
    else:
        return None
    def parse(dt_str):
        if dt_str:
            try:
                return datetime.fromisoformat(dt_str.rstrip("Z"))
            except ValueError:
                pass
        return None
    i_dt = parse(interactive)
    ni_dt = parse(non_interactive)
    if i_dt and ni_dt:
        return max(i_dt, ni_dt)
    return i_dt or ni_dt

def formatted_stale(users):
    formatted = []
    for user in users:
        activity = user.get("signInActivity", {})
        formatted.append({
            "Display Name": user.get("displayName", ""),
            "User Principal Name": user.get("userPrincipalName", ""),
            "Employee ID": user.get("employeeId", ""),
            "Department": user.get("department", ""),
            "Office Location": user.get("officeLocation", ""),
            "Last Successful Sign-In": activity.get("lastSuccessfulSignInDateTime", ""),
            "Last Non-Interactive Sign-In": activity.get("lastNonInteractiveSignInDateTime", ""),
            "Has License": user.get("hasLicense", False),
        })
    return formatted

def has_license(user_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/licenseDetails"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return bool(data.get("value")) 
    return False
if __name__ == "__main__":
    header = {'Authorization': f'Bearer {get_access_token(["https://graph.microsoft.com/.default"])}'}
    users = get_accounts(header)
    stale = stale_accounts(users, header)
    formatted = formatted_stale(stale)
    export_to_xlsx(formatted, r"C:\Users\skylar.eade\Desktop\Python\ms-apis\excel\stale_accounts.xlsx", "Stale Accounts")
        
