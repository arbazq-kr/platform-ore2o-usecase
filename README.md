# platform-ore2o-usecase
This is the repo to maintain use cases code.

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
UPDATED: Two-layer matching now AI-based using sentence-transformers:
Layer 1 → Semantic similarity on short_description (broad match, threshold 0.85)
Layer 2 → Semantic similarity on description (exact match, threshold 0.90)
9o
UPDATED: Alert-based filtering now done at API level via contact_type=event
         Removed is_alert_based() and filter_alert_incidents() as no longer needed
"""