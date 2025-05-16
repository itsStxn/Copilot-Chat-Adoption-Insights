from .custom_types import FINDER
from .imports import os
from .env import (
    URL_SHAREPOINT,
    URL_POWERBI_SHOW,
    URL_POWERBI_EXPORT
)


EDGE_USER_DATA_DIR = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data")
URLS = {
    "outlook": "https://outlook.office.com/mail/inbox",
    "sharepoint": URL_SHAREPOINT,
    "powerbi export": URL_POWERBI_EXPORT,
    "powerbi show": URL_POWERBI_SHOW
}
POWERBI_DOM: FINDER = {
    #? Western Europe overview
    "we mau": (".pivotTable", "Market"),

    #? Yesterday's progress
    "snapshot": (".tableExContainer", None),
    "snapshot container": ("transform", "Account Details (Managed Account Only)"),
    "loader": (".powerbi-spinner", "shown"),
    
    #? Filter area
    "dropdown button": ("[role='combobox'][aria-label='*attr*']", None),
    "dropdown": ("[role='listbox'][aria-label='*attr*']", None),
    "option": ("[role='option'][title='*attr*']", None),
    "option outline": (".setFocusRing", None),
    
    #? Share url section
    "share": ("button[title='Share']", None),
    "url": ("button[data-unique-id='CopyLinkButton']", None),
    "copy url": ("input[id='copyInputBox']", None),
    "close url": ("button[data-unique-id='SharingDialogCloseBtn']", None),

    #? Inside the accounts table
    "top": (".top-viewport", None),
    "mid": (".mid-viewport", None),
    "row": (".row", None),
    "index": ("row-index", None)
}
SHAREPOINT_DOM: FINDER = {
    #? Files
    "snapshots": ("[data-id='heroField']", "Table Data.txt"),
    "V-teams": ("[data-id='heroField']", "V-teams.txt"),
    "MW SSPs": ("[data-id='heroField']", "AE v-teams.txt"),
    "AE mail": ("[data-id='heroField']", "Mail structure - AE.txt"),
    "excluded mail": ("[data-id='heroField']", "Mail structure - Excluded.txt"),
    "excluded": ("[data-id='heroField']", "Excluded from CCs.txt"),
    "skip": ("[data-id='heroField']", "Skip.txt"),
    "filter accounts": ("[data-id='heroField']", "Filter accounts.txt"),
 
    #? Row structure
    "rows": (".view-lines", None),
    "row": (".view-line", None),
    "row count": (".margin-view-overlays", None),
    
    #? Text area
    "editor": (".editor-scrollable", None),
    
    #? Commands & notifications
    "save": ("button[id='saveCommand']", None),
    "close": ("button[id='closeCommand']", None),
    "popup": ("[role='alert']", "Saved")
}
OUTLOOK_DOM: FINDER = {
    #? Commands
    "new mail": ("button", "New mail"),
    "reply": ("button[role='menuitem'][aria-label='Reply']", None),
    
    #? Text area
    "reader": ("[data-app-section='MailReadCompose']", None),
    "subject": ("input[aria-label='Subject']", None),
    "field": ("[role='textbox']", None),
    "send": ("button[aria-label='Send']", None),
    
    #? Search bar and message
    "search": ("input[id='topSearchInput']", None),
    "message": ("[role='option'] [role='group']", None)
}
