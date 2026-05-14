"""
DEV NOTES: ARBAZ QURESHI (TTC4624)
This is a test script, which fetches data from ServiceNow api,
creates dataframe of incidents,
Groups them based on first 50 cleaned characters of short description (clean_description),
before linking asks user permission, and links if yes.
FIXED: Now correctly identifies parent by checking if parent exists within group,
and ignores external parents from different infra groups.
FIXED: If oldest incident is linked to external group, skip only that oldest,
then make second oldest the parent regardless of its external link status.
Link all other incidents in group to this second oldest (chosen parent).
ADDED: Two-layer matching:
Layer 1 → clean_short_description based on short_description field (broad match)
Layer 2 → clean_description based on description field (exact match sub-grouping)
"""

import requests
import re
import os
import pandas as pd
from mysql import connector
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════

api_key  = os.getenv("api_token")
url      = "https://kroger.service-now.com/api/now/table/incident"

headers = {
    "Authorization": api_key,
    "Content-Type":  "application/json",
    "Accept":        "application/json"
}

querystring = {
    "sysparm_query": (
        "assignment_groupIN"
        "7cd3576c1b1a4d508611caa3604bcbfa,"
        "2755644847883e54ed0faf4a216d4396,"
        "194f6bb0c3c37290f792031ad00131a4,"
        "dc8c9dfe1b160d10684965f3604bcbea"
        "^active=true"
        "^stateIN2,4,5"
        "^ORDERBYDESCsys_created_on"
    ),
    "sysparm_fields": (
        "number,sys_id,active,state,"
        "short_description,description,cmdb_ci,u_ci_name,"
        "u_missing_ci,parent_incident,"
        "child_incidents,assignment_group,"
        "assigned_to,sys_created_on,work_notes,"
        "contact_type,u_secure_type"
    ),
    "sysparm_limit":                 "100",
    "sysparm_exclude_reference_link":"true"
}

columns = [
    "number", "sys_id", "active",
    "state", "configuration_item", "u_ci_name",
    "u_missing_ci", "parent_incident", "child_incidents",
    "assignment_group", "assigned_to",
    "short_description", "clean_short_description",
    "description", "clean_description",
    "work_notes", "sys_created_on",
    "contact_type", "u_secure_type"
]


# ══════════════════════════════════════════════════════════
#  FUNCTION: CLEAN SHORT DESCRIPTION
#  Used for Layer 1 broad matching
# ══════════════════════════════════════════════════════════

def clean_short_description(short_description):
    """
    Clean short_description:
    1. Remove URLs
    2. Remove special characters
    3. Strip extra whitespace
    4. Return full cleaned text (no truncation)
    Used for Layer 1 broad group matching
    """
    if not short_description or short_description == "":
        return None

    text = re.sub(r'http[s]?://\S+', '', short_description)
    text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    return text if text else None


# ══════════════════════════════════════════════════════════
#  FUNCTION: CLEAN DESCRIPTION
#  Used for Layer 2 exact matching
# ══════════════════════════════════════════════════════════

def clean_description(description):
    """
    Clean description:
    1. Remove URLs
    2. Remove special characters
    3. Strip extra whitespace
    4. Return full cleaned text (no truncation)
    Used for Layer 2 exact sub-group matching
    """
    if not description or description == "":
        return None

    text = re.sub(r'http[s]?://\S+', '', description)
    text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    return text if text else None


# ══════════════════════════════════════════════════════════
#  FUNCTION: IS ALERT BASED INCIDENT
# ══════════════════════════════════════════════════════════

def is_alert_based(contact_type, u_secure_type):
    contact = str(contact_type).strip().lower() if contact_type else ""
    secure  = str(u_secure_type).strip()        if u_secure_type else ""
    is_event        = contact == "event"
    is_secure_empty = secure == "" or secure.upper() == "NONE"
    if is_event and is_secure_empty:
        return True
    return False


# ══════════════════════════════════════════════════════════
#  FUNCTION: IS ALREADY A CHILD
# ══════════════════════════════════════════════════════════

def is_already_child(parent_incident_field):
    """
    Returns True if incident has a valid parent_incident set
    """
    if not parent_incident_field:
        return False
    if parent_incident_field == '':
        return False
    if parent_incident_field == {}:
        return False
    if isinstance(parent_incident_field, dict):
        value = parent_incident_field.get('value', '')
        return value != ''
    if str(parent_incident_field).strip() == '':
        return False
    return True


# ══════════════════════════════════════════════════════════
#  FUNCTION: GET PARENT SYS ID
#  Returns parent sys_id string if it exists, else None
# ══════════════════════════════════════════════════════════

def get_parent_sys_id(parent_incident_field):
    """
    Extracts and returns parent sys_id string from
    parent_incident field regardless of its type.
    Returns None if no valid parent found.
    """
    if not parent_incident_field:
        return None
    if parent_incident_field == '':
        return None
    if parent_incident_field == {}:
        return None
    if isinstance(parent_incident_field, dict):
        value = parent_incident_field.get('value', '').strip()
        return value if value else None
    value = str(parent_incident_field).strip()
    return value if value else None


# ══════════════════════════════════════════════════════════
#  FUNCTION: LINK CHILD TO PARENT
# ══════════════════════════════════════════════════════════

def link_child_to_parent(child_sys_id, parent_sys_id, child_number):
    """
    Links a child incident to parent incident
    via PATCH request on parent_incident field
    """
    patch_url = f"https://kroger.service-now.com/api/now/table/incident/{child_sys_id}"

    payload  = {"parent_incident": parent_sys_id}
    response = requests.patch(patch_url, headers=headers, json=payload)

    if response.status_code == 200:
        print(f"   ✅ Linked {child_number} → parent successfully!")
        return True
    else:
        print(f"   ❌ Failed to link {child_number}: {response.status_code}")
        print(f"      {response.text}")
        return False


# ══════════════════════════════════════════════════════════
#  FUNCTION: ASK USER PERMISSION
# ══════════════════════════════════════════════════════════

def ask_permission(question):
    """
    Ask user Y/N question
    Returns True if Yes, False if No
    Keeps asking until valid input given
    """
    while True:
        answer = input(f"\n   {question} (Y/N): ").strip().upper()

        if answer == 'Y':
            return True
        elif answer == 'N':
            return False
        else:
            print(f"   ⚠️  Invalid input '{answer}'. Please enter Y or N only.")


# ══════════════════════════════════════════════════════════
#  FUNCTION: STEP 1 → FETCH FROM SERVICENOW
# ══════════════════════════════════════════════════════════

def fetch_incidents():
    """
    Fetches incidents from ServiceNow API
    Returns raw list of incident dicts
    """
    print(f"\n{'='*65}")
    print(f"   🚀 INCIDENT PARENT-CHILD LINKER")
    print(f"{'='*65}")
    print(f"\n🔄 Fetching incidents from ServiceNow...")

    response = requests.get(url, headers=headers, params=querystring)
    data     = response.json()
    datax    = data['result']

    print(f"✅ Fetched {len(datax)} incidents")

    return datax


# ══════════════════════════════════════════════════════════
#  FUNCTION: STEP 2 → BUILD DATAFRAME
# ══════════════════════════════════════════════════════════

def build_dataframe(datax):
    """
    Builds pandas DataFrame from raw incident list
    Adds clean_short_description and clean_description columns
    Returns DataFrame
    """
    rows = []

    for inc in datax:
        raw_short_description = inc.get('short_description', '')
        raw_description       = inc.get('description',       '')

        active_row = [
            inc.get('number',           ''),
            inc.get('sys_id',           ''),
            inc.get('active',           ''),
            inc.get('state',            ''),
            inc.get('cmdb_ci',          ''),
            inc.get('u_ci_name',        ''),
            inc.get('u_missing_ci',     ''),
            inc.get('parent_incident',  ''),
            inc.get('child_incidents',  '0'),
            inc.get('assignment_group', ''),
            inc.get('assigned_to',      ''),
            raw_short_description,
            clean_short_description(raw_short_description),
            raw_description,
            clean_description(raw_description),
            inc.get('work_notes',       ''),
            inc.get('sys_created_on',   None),
            inc.get('contact_type',     ''),
            inc.get('u_secure_type',    '')
        ]
        rows.append(active_row)

    df = pd.DataFrame(rows, columns=columns)

    print(f"📋 DataFrame created with {len(df)} rows")
    print(f"\n📋 Sample:")
    print(df[['number', 'clean_short_description']].head(20))

    return df


# ══════════════════════════════════════════════════════════
#  FUNCTION: STEP 2.5 → FILTER ALERT BASED INCIDENTS ONLY
# ══════════════════════════════════════════════════════════

def filter_alert_incidents(df):
    """
    Filters DataFrame to keep only alert-based incidents
    Prints summary of filtered and skipped incidents
    Returns filtered DataFrame with only alert-based incidents
    """
    print(f"\n{'='*65}")
    print(f"🔍 Filtering Alert-Based Incidents Only...")
    print(f"{'='*65}")

    df['is_alert'] = df.apply(
        lambda row: is_alert_based(
            row['contact_type'],
            row['u_secure_type']
        ),
        axis=1
    )

    alert_df   = df[df['is_alert'] == True].copy()
    skipped_df = df[df['is_alert'] == False].copy()

    print(f"\n📊 Filtering Summary:")
    print(f"   Total Fetched   : {len(df)}")
    print(f"   ✅ Alert Based  : {len(alert_df)}")
    print(f"   🚫 Skipped      : {len(skipped_df)}")

    if not skipped_df.empty:
        print(f"\n   🚫 Skipped Incidents:")
        #print(f"   {'INC Number':<15} {'contact_type':<15} {'u_secure_type'}")
        print(f"   {'-'*50}")
        tempArray = []
        for _, row in skipped_df.iterrows():
            #print(f"   {row['number']:<15} {row['contact_type']:<15} {row['u_secure_type']}")
            tempArray.append(row['number'])
        print(tempArray)

    if not alert_df.empty:
        print(f"\n   ✅ Alert Based Incidents:")
        #print(f"   {'INC Number':<15} {'contact_type':<15} {'u_secure_type'}")
        print(f"   {'-'*50}")
        tempArray = []
        for _, row in alert_df.iterrows():
            #print(f"   {row['number']:<15} {row['contact_type']:<15} {row['u_secure_type']}")
            tempArray.append(row['number'])
        print(tempArray)

    return alert_df


# ══════════════════════════════════════════════════════════
#  FUNCTION: DETERMINE CHOSEN PARENT FOR SUB-GROUP
#  Called inside Layer 2 loop when no in-group parent found
# ══════════════════════════════════════════════════════════

def determine_chosen_parent(group_sorted, oldest, oldest_number, oldest_sys_id):
    """
    Determines which incident becomes the chosen parent
    when no valid in-group parent is found.
    If oldest has external parent → skip it → pick second oldest
    If oldest has no external parent → pick oldest as normal
    Returns tuple:
        chosen_parent_row    → DataFrame row of chosen parent
        chosen_parent_sys_id → sys_id string of chosen parent
        chosen_parent_number → number string of chosen parent
        oldest_parent_sys_id → external parent sys_id of oldest (or None)
    """
    oldest_parent_sys_id = get_parent_sys_id(oldest['parent_incident'])

    if oldest_parent_sys_id is not None:
        # ─── Oldest is externally linked ──────────────────
        # Skip only the oldest
        # Make second oldest the parent no matter what

        if len(group_sorted) < 2:
            # Only 1 incident in group and it's externally linked
            # Nothing to do
            return None, None, None, oldest_parent_sys_id

        # ✅ Pick second oldest as chosen parent
        chosen_parent_row    = group_sorted.iloc[1]
        chosen_parent_sys_id = chosen_parent_row['sys_id']
        chosen_parent_number = chosen_parent_row['number']

        print(f"   ⚠️  Oldest {oldest_number} is linked to external parent")
        print(f"      Skipping oldest as parent candidate")
        print(f"   👑 Will make {chosen_parent_number} the PARENT")
        print(f"      (Second oldest, oldest was externally linked)")

    else:
        # ─── Oldest has no external parent ────────────────
        # Make oldest the chosen parent as normal
        chosen_parent_row    = oldest
        chosen_parent_sys_id = oldest_sys_id
        chosen_parent_number = oldest_number

        print(f"   👑 Will make {chosen_parent_number} the PARENT")
        print(f"      (No valid parent found within this group)")

    return (
        chosen_parent_row,
        chosen_parent_sys_id,
        chosen_parent_number,
        oldest_parent_sys_id
    )


# ══════════════════════════════════════════════════════════
#  FUNCTION: SHOW LINKING PLAN FOR SUB-GROUP
#  Prints what will be linked before asking permission
# ══════════════════════════════════════════════════════════

def show_linking_plan(
    group_sorted,
    existing_parent_sys_id,
    existing_parent_number,
    chosen_parent_sys_id,
    chosen_parent_number,
    oldest_parent_sys_id,
    oldest_sys_id,
    oldest_number
):
    """
    Prints the linking plan for a sub-group.
    Shows which incidents will be linked, re-linked or skipped.
    Returns unlinked_count (used to skip group if nothing to do)
    """
    # ─── Show linking plan ─────────────────────────────────
    print(f"\n   📋 LINKING PLAN:")

    if existing_parent_sys_id and existing_parent_sys_id != '':
        # ══════════════════════════════════════════════════
        #  VALID IN-GROUP PARENT FOUND
        #  Link all unlinked incidents to this parent
        # ══════════════════════════════════════════════════
        print(f"   ⚠️  Parent already exists: {existing_parent_number}")
        print(f"   Will link only UNLINKED incidents to existing parent")

        unlinked_count = 0
        for _, row in group_sorted.iterrows():
            parent_field = row['parent_incident']

            current_parent = ""
            if parent_field and parent_field != '' and parent_field != {}:
                current_parent = (
                    parent_field.get('value', '')
                    if isinstance(parent_field, dict)
                    else str(parent_field).strip()
                )

            if current_parent == existing_parent_sys_id:
                continue

            if row['sys_id'] == existing_parent_sys_id:
                continue

            print(f"   → Will link: {row['number']}")
            unlinked_count += 1

        return unlinked_count

    else:
        # ══════════════════════════════════════════════════
        #  NO VALID IN-GROUP PARENT FOUND
        #  Check if oldest has external parent:
        #  → YES: skip ONLY the oldest (1 incident skipped)
        #         make second oldest the chosen parent
        #         regardless of second oldest external link
        #  → NO : make oldest the chosen parent as normal
        #  Then link ALL others to chosen parent
        #  even if they are linked to external group
        # ══════════════════════════════════════════════════

        # ─── Show who will be linked as children ──────────
        for _, child in group_sorted.iterrows():

            # Skip the chosen parent itself
            if child['sys_id'] == chosen_parent_sys_id:
                continue

            # Skip the oldest if it was externally linked
            # (it stays linked to its external parent)
            if (oldest_parent_sys_id is not None
                    and child['sys_id'] == oldest_sys_id):
                print(f"   ⏩ Skipping {child['number']}")
                print(f"      → stays linked to external parent: {oldest_parent_sys_id}")
                continue

            parent_field      = child['parent_incident']
            parent_sys_id_val = get_parent_sys_id(parent_field)

            if parent_sys_id_val is not None:
                print(f"   → Will re-link: {child['number']} as child")
                print(f"      (currently linked to external: {parent_sys_id_val})")
            else:
                print(f"   → Will link: {child['number']} as child")

        # unlinked_count not needed in this branch
        # return -1 as sentinel so caller knows plan was shown
        return -1


# ══════════════════════════════════════════════════════════
#  FUNCTION: EXECUTE LINKING FOR SUB-GROUP
#  Performs actual PATCH calls after user approves
# ══════════════════════════════════════════════════════════

def execute_linking(
    group_sorted,
    existing_parent_sys_id,
    existing_parent_number,
    chosen_parent_sys_id,
    chosen_parent_number,
    oldest_parent_sys_id,
    oldest_sys_id
):
    """
    Executes the actual linking of incidents.
    Handles two cases:
    Case 1 → existing in-group parent found → link unlinked to it
    Case 2 → no in-group parent → link all to chosen parent
    Returns tuple: (linked, skipped, failed) counts
    """
    linked  = 0
    skipped = 0
    failed  = 0

    # ══════════════════════════════════════════════════════
    #  CASE 1: Parent Already Exists (within group)
    # ══════════════════════════════════════════════════════

    if existing_parent_sys_id and existing_parent_sys_id != '':
        print(f"\n   ⚠️  Using existing parent: {existing_parent_number}")

        for _, row in group_sorted.iterrows():
            parent_field = row['parent_incident']

            current_parent = ""
            if parent_field and parent_field != '' and parent_field != {}:
                current_parent = (
                    parent_field.get('value', '')
                    if isinstance(parent_field, dict)
                    else str(parent_field).strip()
                )

            # Skip if already linked to this parent
            if current_parent == existing_parent_sys_id:
                print(f"   ⏭️  Skipped {row['number']} → already linked")
                skipped += 1
                continue

            # Skip if IS the parent
            if row['sys_id'] == existing_parent_sys_id:
                print(f"   👑 Skipped {row['number']} → IS the parent")
                skipped += 1
                continue

            success = link_child_to_parent(
                child_sys_id  = row['sys_id'],
                parent_sys_id = existing_parent_sys_id,
                child_number  = row['number']
            )

            if success:
                linked += 1
            else:
                failed += 1

    # ══════════════════════════════════════════════════════
    #  CASE 2: No In-Group Parent Found
    #  → Use chosen_parent
    #  → Skip the externally linked oldest (only 1 skip)
    #  → Link ALL others to chosen parent
    #    even if they have external parent links
    # ══════════════════════════════════════════════════════

    else:
        print(f"\n   👑 Making {chosen_parent_number} the PARENT")

        for _, child in group_sorted.iterrows():

            # Skip the chosen parent itself
            if child['sys_id'] == chosen_parent_sys_id:
                print(f"   👑 Skipped {child['number']} → IS the parent")
                skipped += 1
                continue

            # Skip the oldest if it was externally linked
            # (leave it linked to its external parent, do not touch)
            if (oldest_parent_sys_id is not None
                    and child['sys_id'] == oldest_sys_id):
                print(f"   ⏭️  Skipped {child['number']} → externally linked, leaving as-is")
                skipped += 1
                continue

            parent_field = child['parent_incident']

            current_parent = ""
            if parent_field and parent_field != '' and parent_field != {}:
                current_parent = (
                    parent_field.get('value', '')
                    if isinstance(parent_field, dict)
                    else str(parent_field).strip()
                )

            # Skip if already linked to chosen parent
            if current_parent == chosen_parent_sys_id:
                print(f"   ⏭️  Skipped {child['number']} → already linked")
                skipped += 1
                continue

            # Link to chosen parent
            # (even if currently linked to external parent)
            success = link_child_to_parent(
                child_sys_id  = child['sys_id'],
                parent_sys_id = chosen_parent_sys_id,
                child_number  = child['number']
            )

            if success:
                linked += 1
            else:
                failed += 1

    return linked, skipped, failed


# ══════════════════════════════════════════════════════════
#  FUNCTION: PROCESS LAYER 2 SUB-GROUP
#  Handles one sub-group from Layer 2 matching
# ══════════════════════════════════════════════════════════

def process_subgroup(group, current_l2_group, total_l2_groups):
    """
    Processes a single Layer 2 sub-group:
    1. Sorts group oldest first
    2. Prints sub-group info
    3. Checks for existing in-group parent
    4. Determines chosen parent if no in-group parent
    5. Shows linking plan
    6. Asks user permission
    7. Executes linking
    8. Prints summary
    """
    # ─── Convert and Sort ─────────────────────────────────
    group = group.copy()
    group['sys_created_on'] = pd.to_datetime(group['sys_created_on'])
    group_sorted  = group.sort_values('sys_created_on', ascending=True)
    oldest        = group_sorted.iloc[0]
    oldest_number = oldest['number']
    oldest_sys_id = oldest['sys_id']
    oldest_date   = oldest['sys_created_on']
    children      = group_sorted[group_sorted['number'] != oldest_number]

    # ─── Collect all sys_ids in this group ────────────────
    group_sys_ids = set(group_sorted['sys_id'].tolist())

    # ─── Print Sub-Group Info ──────────────────────────────
    print(f"\n{'─'*65}")
    print(f"   📌 Sub-Group {current_l2_group}/{total_l2_groups}")
    print(f"   Description : '{str(oldest['clean_description'])[:80]}...'")
    print(f"   Total Count : {len(group)} incidents")
    print(f"   {'INC Number':<15} {'State':<8} {'Created On':<25} {'Role'}")
    print(f"   {'-'*65}")

    for _, row in group_sorted.iterrows():
        role = "⭐ PARENT (Oldest)" if row['number'] == oldest_number else "👶 Child"
        print(f"   {row['number']:<15} {row['state']:<8} {str(row['sys_created_on']):<25} {role}")

    # ─── Check Existing Parent (FIXED) ────────────────────
    # Only treat as existing parent if parent sys_id
    # belongs to an incident WITHIN this group.
    # If parent is from a DIFFERENT group → ignore it.
    existing_parent_sys_id = None
    existing_parent_number = None

    for _, row in group_sorted.iterrows():
        parent_field = row['parent_incident']

        if parent_field and parent_field != '' and parent_field != {}:
            raw_parent_sys_id = (
                parent_field.get('value', '')
                if isinstance(parent_field, dict)
                else str(parent_field).strip()
            )

            if raw_parent_sys_id in group_sys_ids:
                # ✅ Parent is within our group → valid
                existing_parent_sys_id = raw_parent_sys_id
                existing_parent_number = row['number']
                print(f"\n   🔍 Found parent within group: {existing_parent_number}")
                break
            else:
                # ⚠️ Parent is outside our group → ignore
                print(f"\n   ⚠️  {row['number']} has parent in DIFFERENT group")
                print(f"      Parent sys_id : {raw_parent_sys_id}")
                print(f"      Ignoring this external parent...")

    # ─── Determine chosen parent if no in-group parent ────
    chosen_parent_row    = None
    chosen_parent_sys_id = None
    chosen_parent_number = None
    oldest_parent_sys_id = None

    if not existing_parent_sys_id:
        (
            chosen_parent_row,
            chosen_parent_sys_id,
            chosen_parent_number,
            oldest_parent_sys_id
        ) = determine_chosen_parent(
            group_sorted,
            oldest,
            oldest_number,
            oldest_sys_id
        )

        if chosen_parent_row is None:
            # Only 1 incident and it's externally linked
            # Nothing to do
            print(f"   ⚠️  Only 1 incident and it has external parent!")
            print(f"      Cannot determine parent. Skipping group...")
            print(f"{'─'*65}")
            return

    # ─── Show linking plan ─────────────────────────────────
    unlinked_count = show_linking_plan(
        group_sorted         = group_sorted,
        existing_parent_sys_id = existing_parent_sys_id,
        existing_parent_number = existing_parent_number,
        chosen_parent_sys_id   = chosen_parent_sys_id,
        chosen_parent_number   = chosen_parent_number,
        oldest_parent_sys_id   = oldest_parent_sys_id,
        oldest_sys_id          = oldest_sys_id,
        oldest_number          = oldest_number
    )

    # ─── If existing parent and nothing to link → skip ────
    if existing_parent_sys_id and unlinked_count == 0:
        print(f"   ✅ All incidents already linked! Nothing to do.")
        print(f"{'─'*65}")
        return

    # ══════════════════════════════════════════════════════
    #  ASK USER PERMISSION
    # ══════════════════════════════════════════════════════

    user_approved = ask_permission(
        f"Do you want to link these {len(group)} incidents?"
    )

    if not user_approved:
        print(f"\n   🚫 SKIPPED by user → Sub-Group")
        print(f"{'─'*65}")
        return

    print(f"\n   ✅ User approved! Proceeding with linking...")
    print(f"   {'-'*50}")

    # ─── Execute linking ───────────────────────────────────
    linked, skipped, failed = execute_linking(
        group_sorted           = group_sorted,
        existing_parent_sys_id = existing_parent_sys_id,
        existing_parent_number = existing_parent_number,
        chosen_parent_sys_id   = chosen_parent_sys_id,
        chosen_parent_number   = chosen_parent_number,
        oldest_parent_sys_id   = oldest_parent_sys_id,
        oldest_sys_id          = oldest_sys_id
    )

    # ─── Summary for this sub-group ───────────────────────
    print(f"\n   📊 Sub-Group Summary:")
    print(f"   👑 Parent  : {existing_parent_number if existing_parent_sys_id else chosen_parent_number}")
    print(f"   ✅ Linked  : {linked}")
    print(f"   ⏭️  Skipped : {skipped}")
    print(f"   ❌ Failed  : {failed}")
    print(f"{'─'*65}")


# ══════════════════════════════════════════════════════════
#  FUNCTION: STEP 3 → FIND SIMILAR INCIDENTS + LINK
#  TWO LAYER MATCHING:
#  Layer 1 → clean_short_description (broad match)
#  Layer 2 → clean_description       (exact match)
# ══════════════════════════════════════════════════════════

def find_and_link_similar_incidents(df):
    """
    Finds similar incidents using two-layer matching
    and links them as parent-child in ServiceNow.
    Layer 1 → groups by clean_short_description (broad match)
    Layer 2 → sub-groups by clean_description (exact match)
    """
    print(f"\n{'='*65}")
    print(f"🔍 Finding Similar Incidents (Two-Layer Matching)...")
    print(f"{'='*65}")

    # ══════════════════════════════════════════════════════
    #  LAYER 1: Group by clean_short_description
    # ══════════════════════════════════════════════════════

    similar_layer1 = df.groupby('clean_short_description').filter(lambda x: len(x) > 1)

    if similar_layer1.empty:
        print("✅ No similar incidents found. Nothing to link!")
        return

    total_l1_groups  = similar_layer1.groupby('clean_short_description').ngroups
    current_l1_group = 0

    print(f"⚠️  Found {total_l1_groups} broad similar group(s) from Layer 1!")

    for short_desc, l1_group in similar_layer1.groupby('clean_short_description'):

        current_l1_group += 1

        print(f"\n{'-'*65}")
        print(f"🔵  Group {current_l1_group}/{total_l1_groups} Layer 1")
        print(f"   Short Desc : '{short_desc}'")
        print(f"   Total INCs : {len(l1_group)}")

        # ══════════════════════════════════════════════════
        #  LAYER 2: Sub-group by clean_description
        # ══════════════════════════════════════════════════

        # ─── Safety net: filter None clean_description ────
        # (Edge case: description had only URL → cleaned to None)
        l1_with_desc    = l1_group[l1_group['clean_description'].notna()].copy()
        l1_without_desc = l1_group[l1_group['clean_description'].isna()].copy()

        if not l1_without_desc.empty:
            print(f"\n   ⚠️  {len(l1_without_desc)} incident(s) have no clean_description:")
            for _, row in l1_without_desc.iterrows():
                print(f"      → {row['number']} (skipped from sub-grouping)")

        # ─── Sub-group by clean_description ───────────────
        similar_layer2 = l1_with_desc.groupby('clean_description').filter(lambda x: len(x) > 1)

        if similar_layer2.empty:
            print(f"\n   ✅ No exact matches in Layer 2")
            print(f"      All incidents have different descriptions")
            print(f"      No linking needed for this group")
            print(f"{'='*65}")
            continue

        total_l2_groups  = similar_layer2.groupby('clean_description').ngroups
        current_l2_group = 0

        print(f"\n   ✅ Found {total_l2_groups} exact match sub-group(s) in Layer 2!")

        for clean_desc, group in similar_layer2.groupby('clean_description'):

            current_l2_group += 1

            # ─── Process this sub-group ───────────────────
            process_subgroup(
                group            = group,
                current_l2_group = current_l2_group,
                total_l2_groups  = total_l2_groups
            )

        print(f"{'='*65}")


# ══════════════════════════════════════════════════════════
#  MAIN FUNCTION
# ══════════════════════════════════════════════════════════

def main():
    """
    Main entry point.
    Orchestrates all steps:
    1. Fetch incidents from ServiceNow
    2. Build DataFrame
    3. Filter alert-based incidents only
    4. Find similar incidents and link parent-child
    """

    # ─── STEP 1: Fetch from ServiceNow ────────────────────
    datax = fetch_incidents()

    # ─── STEP 2: Build DataFrame ──────────────────────────
    df = build_dataframe(datax)

    # ─── STEP 2.5: Filter Alert Based Incidents Only ──────
    df = filter_alert_incidents(df)

    if df.empty:
        print(f"\n✅ No alert-based incidents found. Nothing to process!")
        return

    print(f"\n✅ Processing {len(df)} alert-based incidents...")

    # ─── STEP 3: Find Similar + Link Parent-Child ─────────
    find_and_link_similar_incidents(df)

    # ══════════════════════════════════════════════════════
    #  DONE
    # ══════════════════════════════════════════════════════

    print(f"\n{'='*65}")
    print(f"   ✅ ALL DONE!")
    print(f"   → Incidents fetched from ServiceNow")
    print(f"   → Similar incidents identified")
    print(f"   → Parent-Child links created in ServiceNow")
    print(f"{'='*65}")


# ══════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()