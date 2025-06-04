EMPTY_CELLS_REGEX = r'^\s*$'
DATE_FORMAT = "%#m/%#d/%Y"
COLS = {
  "adoption %": "Copilot Chat MAU (Paid +UnPaid)/TAM",
  "incremental": "Copilot Chat H2 Incremental MAU",
  "adoption": "Copilot Chat MAU (Paid +UnPaid)",
  "account": "TopParent",
  "tam": "Total TAM",
  "date": "Date",
  "ae": "AE"
}
COLTYPES = {
  "text": ["TopParent", "AE"],
  "avg": [COLS["adoption %"]],
  "%": [COLS["adoption %"]],
}
