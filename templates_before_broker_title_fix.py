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
    "S1": {
        "name": "S1",
        "label": "S1",
        "keywords": [
            "S-1 CORPORATION",
            "GENERAL INFORMATION",
            "SPACE AVAILABILITY & RENT",
            "시흥 데콘 스마트물류센터",
        ],
        "GI_FIELDS": [
            "소재지",
            "대지면적",
            "연면적",
            "주용도",
            "규모",
            "인근IC",
            "건폐율/용적률",
            "구조",
            "층고",
            "준공년도",
        ],
        "GI_BOX": {"x0": 0.026, "x1": 0.287, "y0": 0.548, "y1": 0.955},
        "GI_VALUE_X": 0.110,
        "TITLE_BOX": {"x0": 0.025, "x1": 0.610, "y0": 0.018, "y1": 0.115},
        "PHOTO_BOX": {"x0": 0.026, "x1": 0.334, "y0": 0.194, "y1": 0.508},
        "SPACE_BOX": {"x0": 0.512, "x1": 0.982, "y0": 0.525, "y1": 0.950},
        "SPACE_PAD": {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0},
        "SPEC_BOX": {"x0": 0.296, "x1": 0.510, "y0": 0.548, "y1": 0.950},
        "OCR_LANG": "kor+eng",
    },

    "Mateplus": {
    "name": "Mateplus",
    "label": "MatePlus",
    "keywords": [
        "MatePlus",
        "GENERAL INFORMATION",
        "SPACE AVAILABILITY & RENT",
        "PERSPECTIVE VIEW",
    ],
    "GI_FIELDS": [
        "소재지",
        "연면적",
        "주용도",
        "규모",
        "인근IC",
        "건폐율/용적률",
        "준공년도",
    ],
    "GI_BOX": {
        "x0": 0.021,
        "x1": 0.250,
        "y0": 0.585,
        "y1": 0.910,
    },
    "GI_VALUE_X": 0.106,

    # MatePlus titles are obtained from the INDEX table
    # by matching each property's address.
    "TITLE_SOURCE": "index",

    # Do not include "INDEX" because it is graphical,
    # not selectable PDF text.
    "INDEX_PAGE_REQUIRED_TOKENS": [
        "물류센터명",
        "소재지",
        "입주시기",
    ],

    "PHOTO_BOX": {
        "x0": 0.021,
        "x1": 0.325,
        "y0": 0.185,
        "y1": 0.515,
    },
    "SPACE_BOX": {
        "x0": 0.492,
        "x1": 0.955,
        "y0": 0.535,
        "y1": 0.910,
    },
    "SPACE_PAD": {
        "left": 0.0,
        "right": 0.0,
        "top": 0.0,
        "bottom": 0.0,
    },
    "SPEC_BOX": {
        "x0": 0.335,
        "x1": 0.482,
        "y0": 0.850,
        "y1": 0.910,
    },
    "OCR_LANG": "kor+eng",
    "PAGE_REQUIRED_TOKENS": [
    "PERSPECTIVE VIEW",
    "GENERAL INFORMATION",
    "SPACE AVAILABILITY & RENT",
    "소재지",
    "연면적",
],
"PAGE_REQUIRED_MIN": 4,
},
   
    "ADF": {
        "name": "ADF",
        "label": "ADF Asset Management",

        "keywords": [
            "ADF ASSET MANAGEMENT",
            "DC 밀양 물류센터",
            "임대 정보",
            "기타사항",
            "경상남도 밀양시",
        ],

        # Only the structured PROFILE / 임대 정보 page is valid.
        "PAGE_REQUIRED_TOKENS": [
            "임대 정보",
            "소재지",
            "대지면적",
            "연면적",
            "기타사항",
        ],
        "PAGE_REQUIRED_MIN": 4,

        # Visual row order of the right-side information table.
        "GI_FIELDS": [
            "소재지",
            "대지면적",
            "연면적",
            "층수",
            "기둥간격",
            "야드",
            "층고",
            "준공년도",
        ],

        # Right-side general-information table.
        "GI_BOX": {
            "x0": 0.374,
            "x1": 0.720,
            "y0": 0.193,
            "y1": 0.535,
        },

        # Divider between labels and values.
        "GI_VALUE_X": 0.460,

        # Blue title box containing "DC 밀양 물류센터".
        "TITLE_BOX": {
            "x0": 0.018,
            "x1": 0.185,
            "y0": 0.112,
            "y1": 0.160,
        },

        "TITLE_REQUIRED_TOKENS": [
            "물류센터",
        ],

        # Building photograph from the main profile page only.
        "PHOTO_BOX": {
            "x0": 0.019,
            "x1": 0.360,
            "y0": 0.193,
            "y1": 0.907,
        },

        # ADF has no separate availability table.
        # Use the complete right-side information block as 공실현황.
        "SPACE_BOX": {
            "x0": 0.374,
            "x1": 0.720,
            "y0": 0.193,
            "y1": 0.907,
        },

        "SPACE_PAD": {
            "left": 0.0,
            "right": 0.0,
            "top": 0.0,
            "bottom": 0.0,
        },

        # 기타사항 values only.
        "SPEC_BOX": {
            "x0": 0.470,
            "x1": 0.720,
            "y0": 0.700,
            "y1": 0.900,
        },

        "SPEC_FIELD": "기타",
        "OCR_LANG": "kor+eng",
    },
}
DEFAULT_TEMPLATE = "MOVE"
