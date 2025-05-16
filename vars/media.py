from .custom_types import IMG

IMAGE_CROP_THRESHOLD=220
INCENTIVES_SLIDE_PATH = "media/Copilot Chat FRC.png"
DF2IMG_OPTIONS = {
    "df": None,
    "print_index": False,
    "fig_size": (1500, 500), 
    "tbl_header": dict(align="center", fill_color="lightblue", font=dict(color="black", size=12)),
    "tbl_cells": dict(align="center", font=dict(color="black", size=12))
}
MEDIA: dict[str, IMG] = {}
MEDIA_LABELS = {
  "week progress": 0,
  "accounts": 1
}
