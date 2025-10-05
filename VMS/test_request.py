import requests
import os
from html.parser import HTMLParser


# Minimal HTML parser that finds hidden input fields by name
class InputFinder(HTMLParser):
    def __init__(self, names=("csrfmiddlewaretoken", "_token", "authenticity_token", "csrf_token")):
        super().__init__()
        self.names = set(names)
        self.found = {}  # {name: value}

    def handle_starttag(self, tag, attrs):
        if tag != "input":
            return
        attrs = dict(attrs)
        name = attrs.get("name")
        typ = attrs.get("type", "").lower()
        if name in self.names and (typ in ("hidden", "", None)):
            val = attrs.get("value")
            if val:
                self.found[name] = val

BASE_URL = os.getenv("BASE_URL")
operator_email = os.getenv("OPERATOR_EMAIL_ADDRESS")
operator_password = os.getenv("OPERATOR_PASSWORD")

LOGIN_API = f"{BASE_URL}/login"
CREATE_VISITOR = f"{BASE_URL}/admin/visitors"
GET_VISITOR = f"{BASE_URL}/admin/get-visitors"

creds = {"email": f"{operator_email}", "password": f"{operator_password}"}


with requests.Session() as s:
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (requests)",
        "Accept": "application/json",
    })

    # 1) Authenticate
    r = s.get(LOGIN_API)
    r.raise_for_status()
    
    parser = InputFinder()
    parser.feed(r.text)

    # Try hidden input first
    csrf_field, csrf_value = None, None
    if parser.found:
        csrf_field, csrf_value = next(iter(parser.found.items()))
        creds[csrf_field] = csrf_value
        print("Found CSRF token in login page:", csrf_field, csrf_value)

    # If the server sets auth cookies (e.g., Set-Cookie: session=...; HttpOnly),
    # requests.Session will store them automatically in s.cookies.
    # Some APIs also return a CSRF token in headers/body; add it if needed:
    r = s.post(LOGIN_API, json=creds, timeout=15)
    r.raise_for_status()

    
    parser = InputFinder()
    parser.feed(r.text)

    # Try hidden input first
    csrf_field, csrf_value = None, None
    if parser.found:
        csrf_field, csrf_value = next(iter(parser.found.items()))
        print("Found CSRF token in login page:", csrf_field, csrf_value)

    xsrf_cookie = (
        s.cookies.get("XSRF-TOKEN")
    )
    headers = {}
    if xsrf_cookie:
        headers["X-CSRF-TOKEN"] = csrf_value
        headers["X-XSRF-TOKEN"] = csrf_value

    try:
        resp = s.get(GET_VISITOR, timeout=10)
        print("Status:", resp.status_code)
        print("Response headers:", resp.headers)
        respData = resp.json()
        # print("Body:", respData)
    except requests.RequestException as e:
        print("Request failed:", e)
        
    opusvms_session = s.cookies.get("opusvms_session")


    # Form fields (non-file fields)
    data = {
        "full_name": "test pname",
        "phone": "test number",
        "email": "test@gmail.com",
        "national_identification_no": "test ic",
        "unit_no": "1-1-1",
        "car_park_lot": "p1",
        "booking_source": "test",
    }
    data[csrf_field] = csrf_value
    createResp = {}
    try:
        createResp = s.post(CREATE_VISITOR, headers=headers,data=data, timeout=30)
        print("Status:", createResp.status_code)
        print("Response headers:", createResp.headers)
        print("Body:", createResp)
    except requests.RequestException as e:
        print("Request failed:", e)