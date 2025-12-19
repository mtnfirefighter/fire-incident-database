# app_caller_SAVE_FIX.py
# Replacement app.py snippet – ensures CallerName and CallerPhone
# are saved into the incident master record used by HTML print.

# NOTE: This file assumes caller_name and caller_phone are already
# collected in Incident Details.

def _apply_caller_to_incident(incident, caller_name, caller_phone):
    incident["CallerName"] = caller_name
    incident["CallerPhone"] = caller_phone
    return incident