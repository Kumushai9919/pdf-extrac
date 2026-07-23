from __future__ import annotations

TEMPLATES = {
    "MOVE": {
        "name": "MOVE",
        "label": "MOVE",
        "keywords": ["MOVE", "General Information", "Space Availability"],
        "GI_BOX": {"x0": 0.060, "x1": 0.302, "y0": 0.596, "y1": 0.880},
        "GI_VALUE_X": 0.133,
        "TITLE_BOX": {"x0": 0.056, "x1": 0.485, "y0": 0.142, "y1": 0.190},
        "PHOTO_BOX": {"x0": 0.057, "x1": 0.306, "y0": 0.224, "y1": 0.572},
        "SPACE_BOX": {"x0": 0.497, "x1": 0.978, "y0": 0.577, "y1": 0.893},
        "SPEC_BOX": {"x0": 0.304, "x1": 0.496, "y0": 0.603, "y1": 0.873},
        "OCR_LANG": "kor+eng",
    },
    "Mapletree Korea": {
        "name": "Mapletree Korea",
        "label": "메이플트리코리아",
        "keywords": [
            "Mapletree",
            "물류센터 개요",
            "임대 가능 면적",
            "임대 시기",
        ],
        "GI_BOX": {"x0": 0.505, "x1": 0.978, "y0": 0.155, "y1": 0.960},
        "GI_VALUE_X": 0.645,
        "TITLE_BOX": {"x0": 0.015, "x1": 0.500, "y0": 0.025, "y1": 0.105},
        "PHOTO_BOX": {"x0": 0.018, "x1": 0.498, "y0": 0.155, "y1": 0.590},
        "SPACE_BOX": {"x0": 0.015, "x1": 0.500, "y0": 0.595, "y1": 0.965},
        "SPEC_BOX": {"x0": 0.645, "x1": 0.978, "y0": 0.885, "y1": 0.960},
        "OCR_LANG": "kor+eng",
    },

    "CBRE": {
        "name": "CBRE",
        "label": "CBRE",
        "keywords": [
            "CBRE",
            "Logistics Center",
            "Building Information",
            "Available for Lease",
            "화성JW 물류센터",
        ],
        "GI_BOX": {
            "x0": 0.400,
            "x1": 0.705,
            "y0": 0.205,
            "y1": 0.585,
        },
        "GI_VALUE_X": 0.510,
        "TITLE_BOX": {
            "x0": 0.020,
            "x1": 0.600,
            "y0": 0.100,
            "y1": 0.180,
        },
        "PHOTO_BOX": {
            "x0": 0.023,
            "x1": 0.385,
            "y0": 0.235,
            "y1": 0.585,
        },
        "SPACE_BOX": {
            "x0": 0.400,
            "x1": 0.963,
            "y0": 0.588,
            "y1": 0.958,
        },
        "SPEC_BOX": {
            "x0": 0.505,
            "x1": 0.705,
            "y0": 0.535,
            "y1": 0.585,
        },
        "OCR_LANG": "kor+eng",
    },
}
DEFAULT_TEMPLATE = "MOVE"
