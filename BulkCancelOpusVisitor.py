from datetime import datetime
import os
import pytz
import requests
import pandas as pd
from html.parser import HTMLParser
import dotenv
dotenv.load_dotenv()

operatorEmailAddress = os.getenv("OPERATOR_EMAIL_ADDRESS")
operatorVmsPassword = os.getenv("OPERATOR_VMS_PASSWORD")
completed_counter = 0
failed_counter = 0

## VMS URLs
VMS_BASE_URL = os.getenv("VMS_BASE_URL")
LOGIN_API = f"{VMS_BASE_URL}/login"
GET_VISITORS_URL = f"{VMS_BASE_URL}/app/guests/get-visitors"
CANCEL_VISITOR_URL = f"{VMS_BASE_URL}/app/guests"

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

# Function to fetch visitors from VMS
def fetch_visitors(session, csrf_value, start=0, length=100):
    """
    Fetch visitors from VMS using the GET endpoint.
    
    Args:
        session: requests.Session with authenticated cookies
        csrf_value: CSRF token value for request
        start: Starting record index (for pagination)
        length: Number of records to fetch
    
    Returns:
        Tuple of (list of visitor objects, total count) or (None, 0) if failed
    """
    params = {
        'draw': 10,
        'columns[0][data]': 'reg_no',
        'columns[0][name]': 'reg_no',
        'columns[0][searchable]': 'true',
        'columns[0][orderable]': 'true',
        'columns[0][search][value]': '',
        'columns[0][search][regex]': 'false',
        'columns[1][data]': 'name',
        'columns[1][name]': 'name',
        'columns[1][searchable]': 'true',
        'columns[1][orderable]': 'true',
        'columns[1][search][value]': '',
        'columns[1][search][regex]': 'false',
        'columns[2][data]': 'phone',
        'columns[2][name]': 'phone',
        'columns[2][searchable]': 'true',
        'columns[2][orderable]': 'true',
        'columns[2][search][value]': '',
        'columns[2][search][regex]': 'false',
        'columns[3][data]': 'national_identification_no',
        'columns[3][name]': 'national_identification_no',
        'columns[3][searchable]': 'true',
        'columns[3][orderable]': 'true',
        'columns[3][search][value]': '',
        'columns[3][search][regex]': 'false',
        'columns[4][data]': 'unit_no',
        'columns[4][name]': 'unit_no',
        'columns[4][searchable]': 'true',
        'columns[4][orderable]': 'true',
        'columns[4][search][value]': '',
        'columns[4][search][regex]': 'false',
        'columns[5][data]': 'planned_checkin',
        'columns[5][name]': 'planned_checkin_at',
        'columns[5][searchable]': 'true',
        'columns[5][orderable]': 'true',
        'columns[5][search][value]': '',
        'columns[5][search][regex]': 'false',
        'columns[6][data]': 'planned_checkout',
        'columns[6][name]': 'planned_checkout_at',
        'columns[6][searchable]': 'true',
        'columns[6][orderable]': 'true',
        'columns[6][search][value]': '',
        'columns[6][search][regex]': 'false',
        'columns[7][data]': 'status',
        'columns[7][name]': 'status',
        'columns[7][searchable]': 'true',
        'columns[7][orderable]': 'true',
        'columns[7][search][value]': '',
        'columns[7][search][regex]': 'false',
        'columns[8][data]': 'created_by',
        'columns[8][name]': 'created_by',
        'columns[8][searchable]': 'true',
        'columns[8][orderable]': 'true',
        'columns[8][search][value]': '',
        'columns[8][search][regex]': 'false',
        'columns[9][data]': 'operator_company',
        'columns[9][name]': 'operator_company',
        'columns[9][searchable]': 'false',
        'columns[9][orderable]': 'true',
        'columns[9][search][value]': '',
        'columns[9][search][regex]': 'false',
        'columns[10][data]': 'action',
        'columns[10][name]': 'action',
        'columns[10][searchable]': 'false',
        'columns[10][orderable]': 'false',
        'columns[10][search][value]': '',
        'columns[10][search][regex]': 'false',
        'order[0][column]': 0,
        'order[0][dir]': 'desc',
        'order[0][name]': 'reg_no',
        'start': start,
        'length': length,
        'search[value]': '',
        'search[regex]': 'false',
        'status_filter': 1,  # Filter by "Pre-registered" status
        'checkin_date_filter': '',
        'checkout_date_filter': '',
        'search_filter': '',
        '_': int(datetime.now().timestamp() * 1000)
    }
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.5',
        'X-CSRF-TOKEN': csrf_value,
        'X-Requested-With': 'XMLHttpRequest',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin'
    }
    
    try:
        print(f"\nFetching visitors from VMS (start={start}, length={length})...")
        response = session.get(
            GET_VISITORS_URL,
            params=params,
            headers=headers,
            timeout=30
        )
        
        if response.status_code == 200:
            data = response.json()
            visitors = data.get('data', [])
            total = data.get('recordsTotal', 0)
            print(f"✓ Fetched {len(visitors)} visitors (Total: {total})")
            return visitors, total
        else:
            print(f"✗ Failed to fetch visitors: Status {response.status_code}")
            return None, 0
            
    except Exception as e:
        print(f"✗ Error fetching visitors: {e}")
        return None, 0

# Function to cancel a single visitor
def cancel_visitor(session, csrf_value, visitor_id, visitor_name, cancel_reason="old guest"):
    """
    Cancel a single visitor by ID.
    
    Args:
        session: requests.Session with authenticated cookies
        csrf_value: CSRF token value for request
        visitor_id: ID of the visitor to cancel
        visitor_name: Name of the visitor (for logging)
        cancel_reason: Reason for cancellation
    
    Returns:
        True if successful, False otherwise
    """
    global completed_counter
    global failed_counter
    
    cancel_url = f"{CANCEL_VISITOR_URL}/{visitor_id}/cancel"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:146.0) Gecko/20100101 Firefox/146.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1'
    }
    
    data = {
        '_token': csrf_value,
        'cancel_reason': cancel_reason
    }
    
    try:
        response = session.post(
            cancel_url,
            data=data,
            headers=headers,
            timeout=15
        )
        
        if response.status_code == 200 or response.status_code == 302:
            print(f"✓ Cancelled visitor ID {visitor_id}: {visitor_name}")
            completed_counter += 1
            return True
        else:
            print(f"✗ Failed to cancel visitor ID {visitor_id}: {visitor_name} (Status: {response.status_code})")
            failed_counter += 1
            return False
            
    except Exception as e:
        print(f"✗ Error cancelling visitor ID {visitor_id}: {visitor_name} - {e}")
        failed_counter += 1
        return False

# Function to batch cancel visitors
def batch_cancel_visitors(session, csrf_value, cancel_reason="old guest", days_threshold=7):
    """
    Fetch all visitors and cancel them in batch if checkin date is more than specified days ago.
    
    Args:
        session: requests.Session with authenticated cookies
        csrf_value: CSRF token value for request
        cancel_reason: Reason for cancellation
        days_threshold: Only cancel if checkin was more than this many days ago (default: 7)
    """
    global completed_counter
    global failed_counter
    
    completed_counter = 0
    failed_counter = 0
    
    # Fetch visitors in batches
    all_visitors = []
    start = 0
    batch_size = 1000
    
    while True:
        visitors, total = fetch_visitors(session, csrf_value, start=start, length=batch_size)
        
        if visitors is None:
            print("\n✗ Failed to fetch visitors. Exiting.")
            return
        
        all_visitors.extend(visitors)
        
        print(f"Collected {len(all_visitors)}/{total} visitors")
        
        # Check if we've fetched all visitors
        if len(all_visitors) >= total or len(visitors) == 0:
            break
        
        start += batch_size
    
    # Filter visitors by checkin date
    current_date = datetime.now(pytz.timezone('Asia/Kuala_Lumpur'))
    visitors_to_cancel = []
    skipped_count = 0
    
    print(f"\n--- Filtering Visitors by Checkin Date ---")
    print(f"Current date: {current_date.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Threshold: More than {days_threshold} days ago\n")
    
    for visitor in all_visitors:
        visitor_id = visitor.get('id')
        visitor_name = visitor.get('name', 'Unknown')
        checkin_date_str = visitor.get('planned_checkin_at')
        
        if not checkin_date_str:
            print(f"⚠ Skipping visitor ID {visitor_id} ({visitor_name}): No check-in date")
            skipped_count += 1
            continue
        
        # Parse the ISO format date: "2025-05-30T15:00:00.000000Z"
        checkin_date = datetime.strptime(checkin_date_str, "%Y-%m-%dT%H:%M:%S.%fZ")
        checkin_date = pytz.timezone('Asia/Kuala_Lumpur').localize(checkin_date)
                    
        # Skip visitors whose checkin date is already in the past
        if checkin_date < current_date:
            skipped_count += 1
            print(f"○ Skipping ID {visitor_id} ({visitor_name}): Already checked in on {checkin_date.strftime('%Y-%m-%d')} (past guest)")
            continue
        
        try:
            # Calculate days difference
            days_diff = (checkin_date - current_date).days
            
            if days_diff > days_threshold:
                visitors_to_cancel.append(visitor)
                print(f"✓ Will cancel ID {visitor_id} ({visitor_name}): Checkin {days_diff} days ago ({checkin_date.strftime('%Y-%m-%d')})")
            else:
                skipped_count += 1
                print(f"○ Skipping ID {visitor_id} ({visitor_name}): Checkin only {days_diff} days ago ({checkin_date.strftime('%Y-%m-%d')})")
        
        except Exception as e:
            print(f"⚠ Skipping visitor ID {visitor_id} ({visitor_name}): Error parsing date - {e}")
            skipped_count += 1
    
    print(f"\n--- Visitors Pending Cancellation ---")
    print(f"Total visitors fetched: {len(all_visitors)}")
    print(f"Visitors to cancel: {len(visitors_to_cancel)}")
    print(f"Visitors skipped: {skipped_count}")
    print(f"Cancel reason: {cancel_reason}\n")

    if len(visitors_to_cancel) == 0:
        print("No visitors meet the cancellation criteria.")
        return

    # Export visitors to CSV before cancelling
    import time
    todayDate = datetime.now().strftime("%Y-%m-%d")
    timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
    csv_filename = f"visitors_to_cancel_{timestamp}.csv"
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs", todayDate, csv_filename)
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)

    rows = []
    for visitor in visitors_to_cancel:
        rows.append({
            "id": visitor.get("id"),
            "name": visitor.get("name", "Unknown"),
            "planned_checkin_at": visitor.get("planned_checkin_at"),
            "planned_checkout_at": visitor.get("planned_checkout_at"),
            "created_at": visitor.get("created_at"),
        })
    df_preview = pd.DataFrame(rows)
    df_preview = df_preview.sort_values(by=["planned_checkin_at", "name"], ascending=[True, True]).reset_index(drop=True)
    df_preview.to_csv(csv_path, index=False)
    print(f"Preview saved to: {csv_path}\n")

    print(df_preview.to_string(index=False))

    confirm_cancel = input(f"\nProceed to cancel the {len(visitors_to_cancel)} visitors listed above? (yes/no): ").strip().lower()
    if confirm_cancel != "yes":
        print("\nCancellation aborted by user. No visitors were cancelled.")
        return

    # Cancel each filtered visitor
    print(f"\n--- Starting Batch Cancellation ---")
    for idx, visitor in enumerate(visitors_to_cancel, 1):
        visitor_id = visitor.get('id')
        visitor_name = visitor.get('name', 'Unknown')
        checkin_date_str = visitor.get('planned_checkin_at')
        checkout_date_str = visitor.get('planned_checkout_at')
        created_at_str = visitor.get('created_at')
        print(f"[{idx}/{len(visitors_to_cancel)}] Cancelling visitor: ID={visitor_id}, Name={visitor_name}, Checkin={checkin_date_str}, Checkout={checkout_date_str}, Created={created_at_str}")
        cancel_visitor(session, csrf_value, visitor_id, visitor_name, cancel_reason)
    
    print(f"\n--- Cancellation Summary ---")
    print(f"Total visitors fetched: {len(all_visitors)}")
    print(f"Visitors to cancel: {len(visitors_to_cancel)}")
    print(f"Successfully cancelled: {completed_counter}")
    print(f"Failed to cancel: {failed_counter}")
    print(f"Skipped (within {days_threshold} days): {skipped_count}")

def login_vms_and_get_token(session):
    print("\nLogging into Opus VMS to get tokens...")
    creds = {"email": f"{operatorEmailAddress}", "password": f"{operatorVmsPassword}"}
    # 1) Authenticate
    r = session.get(LOGIN_API)
    r.raise_for_status()
    
    parser = InputFinder()
    parser.feed(r.text)

    # Try hidden input first
    csrf_field, csrf_value = None, None
    if parser.found:
        csrf_field, csrf_value = next(iter(parser.found.items()))
        creds[csrf_field] = csrf_value
        # print("Found CSRF token in login page:", csrf_field, csrf_value)

    # If the server sets auth cookies (e.g., Set-Cookie: session=...; HttpOnly),
    # requests.Session will store them automatically in s.cookies.
    # Some APIs also return a CSRF token in headers/body; add it if needed:
    r = session.post(LOGIN_API, json=creds, timeout=15)
    r.raise_for_status()

    
    parser = InputFinder()
    parser.feed(r.text)

    # Refresh CSRF token after login
    csrf_field, csrf_value= None, None
    if parser.found:
        field, value = next(iter(parser.found.items()))
        if field == "_token":
            csrf_field, csrf_value = field, value
            # print("Found CSRF token after logged in:", csrf_field, csrf_value)
    return csrf_field, csrf_value


# Main function to fetch and cancel visitors
def main():
    with requests.Session() as s:
        print("=== Bulk Cancel Opus Visitors ===")
        
        # Login to VMS and get CSRF token
        csrf_field, csrf_value = login_vms_and_get_token(session=s)
        if not csrf_field or not csrf_value:
            print("Failed to retrieve CSRF token from VMS. Exiting.")
            return
        
        # Ask for days threshold
        days_input = input("\nEnter days threshold (only cancel if checkin is more than X days ago, default: 7): ").strip()
        try:
            days_threshold = int(days_input) if days_input else 7
        except ValueError:
            print("Invalid input. Using default threshold of 7 days.")
            days_threshold = 7
        
        # Ask for cancel reason
        cancel_reason = input("\nEnter cancellation reason (default: 'old guest'): ").strip()
        if not cancel_reason:
            cancel_reason = "old guest"
        
        # Confirm before proceeding
        confirm = input(f"\nThis will cancel visitors whose checkin date is more than {days_threshold} days ago with reason: '{cancel_reason}'\nContinue? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("\nOperation cancelled by user.")
            return
        
        # Perform batch cancellation
        batch_cancel_visitors(
            session=s,
            csrf_value=csrf_value,
            cancel_reason=cancel_reason,
            days_threshold=days_threshold
        )
        
        print(f'\n=== Final Results ===')
        print(f'Completed: {completed_counter}')
        print(f'Failed: {failed_counter}')
        return


if __name__ == "__main__":
    main()

