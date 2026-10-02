from __future__ import annotations

TEMPLATES = {
    "MOVE": {
        "name": "MOVE",
        "label": "MOVE",
        "keywords": ["MOVE", "General Information", "Space Availability"],
        "GI_BOX": {"x0": 0.060, "x1": 0.302, "y0": 0.596, "y1": 0.880},
        "GI_VALUE_X": 0.133,

        # Wider/taller than before so OCR receives the complete bold headline.
        "TITLE_BOX": {"x0": 0.035, "x1": 0.680, "y0": 0.125, "y1": 0.195},
        "TITLE_REQUIRED_TOKENS": ["물류센터"],
        "TITLE_REQUIRED_TOKENS_OPTIONAL": True,
        "TITLE_REJECT_TOKENS": [
            "Master Of Value Enhancement",
            "Leasing Information",
            "공실 요약",
            "동남권",
        ],
        "TITLE_OCR_PSMS": [7, 6, 11],
        "TITLE_OCR_SCALE": 2,

        # MOVE files can be image-only. Both structural headings must exist,
        # otherwise cover/summary/region-divider pages are skipped.
        "OCR_PAGE_MARKERS": [
            {
                "box": {"x0": 0.030, "x1": 0.470, "y0": 0.550, "y1": 0.670},
                "tokens": ["General Information"],
                "psm": 11,
            },
            {
                "box": {"x0": 0.470, "x1": 0.830, "y0": 0.540, "y1": 0.670},
                "tokens": ["Space Availability"],
                "psm": 11,
            },
        ],
        "OCR_PAGE_REQUIRED_MIN": 2,

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
        "PAGE_REQUIRED_TOKENS": [
            "물류센터 개요",
            "임대 가능 면적",
        ],
        "PAGE_REQUIRED_MIN": 2,
        "GI_BOX": {"x0": 0.505, "x1": 0.978, "y0": 0.155, "y1": 0.960},
        "GI_VALUE_X": 0.645,
        "TITLE_BOX": {"x0": 0.015, "x1": 0.500, "y0": 0.025, "y1": 0.105},
        "TITLE_REQUIRED_TOKENS": ["물류센터"],
        "TITLE_REJECT_TOKENS": [
            "Mapletree Korea",
            "Mapletree",
            "Table of Contents",
            "km",
            "분) 이내",
        ],
        "TITLE_OCR_FALLBACK": True,
        "TITLE_OCR_PSMS": [7, 6, 11],
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
        "GI_BOX": {"x0": 0.400, "x1": 0.705, "y0": 0.205, "y1": 0.585},
        "GI_VALUE_X": 0.510,
        "TITLE_BOX": {"x0": 0.020, "x1": 0.600, "y0": 0.100, "y1": 0.180},
        "TITLE_REQUIRED_TOKENS": ["물류센터"],
        "TITLE_REQUIRED_TOKENS_OPTIONAL": True,
        "TITLE_REJECT_TOKENS": [
            "Logistics Center",
            "For Lease",
            "권역별 임대자산",
            "전국 권역별 물류센터 통합 임대 안내문",
            "Leasing Flyer",
        ],
        "TITLE_OCR_PSMS": [7, 6, 11],
        "TITLE_OCR_SCALE": 2,

        # CBRE test files can also be image-only. Require both table headings.
        "OCR_PAGE_MARKERS": [
            {
                "box": {"x0": 0.380, "x1": 0.730, "y0": 0.180, "y1": 0.280},
                "tokens": ["Building Information"],
                "psm": 11,
            },
            {
                "box": {"x0": 0.380, "x1": 0.780, "y0": 0.550, "y1": 0.660},
                "tokens": ["Available for Lease"],
                "psm": 11,
            },
        ],
        "OCR_PAGE_REQUIRED_MIN": 2,

        "PHOTO_BOX": {"x0": 0.023, "x1": 0.385, "y0": 0.235, "y1": 0.585},
        "SPACE_BOX": {"x0": 0.400, "x1": 0.963, "y0": 0.588, "y1": 0.958},
        "SPEC_BOX": {"x0": 0.505, "x1": 0.705, "y0": 0.535, "y1": 0.585},
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
        "PAGE_REQUIRED_TOKENS": [
            "PERSPECTIVE VIEW",
            "GENERAL INFORMATION",
            "SPACE AVAILABILITY & RENT",
            "소재지",
            "연면적",
        ],
        "PAGE_REQUIRED_MIN": 4,
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
        # The old y1=0.115 crop cut through the lower half of the large S1
        # headline on some pages (for example 평택 하북리), causing blank OCR.
        "TITLE_BOX": {"x0": 0.025, "x1": 0.800, "y0": 0.018, "y1": 0.160},
        # Most S1 titles contain 물류센터, but some valid titles such as
        # 오산 LOGIPOLIS do not. A valid temperature suffix is also accepted.
        "TITLE_REQUIRED_PATTERNS": [
            r"물류센터",
            r"\((?:상온|저온)(?:\s*/\s*(?:상온|저온))?\)",
        ],
        "TITLE_REJECT_TOKENS": [
            "S-1 CORPORATION",
            "PERSPECTIVE VIEW",
            "LOCATION",
            "LAYOUT",
        ],
        "TITLE_OCR_FALLBACK": True,
        "TITLE_OCR_PSMS": [7, 6, 11],
        "TITLE_OCR_SCALE": 1,
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
        "PAGE_REQUIRED_TOKENS": [
            "GENERAL INFORMATION",
            "SPACE AVAILABILITY & RENT",
            "소재지",
            "연면적",
            "준공일",
        ],
        "PAGE_REQUIRED_MIN": 3,
        "GI_FIELDS": [
            "소재지",
            "연면적",
            "주용도",
            "규모",
            "인근IC",
            "건폐율/용적률",
            "준공년도",
        ],
        "GI_BOX": {"x0": 0.021, "x1": 0.250, "y0": 0.585, "y1": 0.910},
        "GI_VALUE_X": 0.106,
        "TITLE_SOURCE": "index",
        "INDEX_PAGE_REQUIRED_TOKENS": [
            "물류센터명",
            "소재지",
            "입주시기",
        ],
        "PHOTO_BOX": {"x0": 0.021, "x1": 0.325, "y0": 0.185, "y1": 0.515},
        "SPACE_BOX": {"x0": 0.492, "x1": 0.955, "y0": 0.535, "y1": 0.910},
        "SPACE_PAD": {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0},
        "SPEC_BOX": {"x0": 0.335, "x1": 0.482, "y0": 0.850, "y1": 0.910},
        "OCR_LANG": "kor+eng",
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
        "PAGE_REQUIRED_TOKENS": [
            "임대 정보",
            "소재지",
            "대지면적",
            "연면적",
            "기타사항",
        ],
        "PAGE_REQUIRED_MIN": 4,
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
        "GI_BOX": {"x0": 0.374, "x1": 0.720, "y0": 0.193, "y1": 0.535},
        "GI_VALUE_X": 0.460,
        "TITLE_BOX": {"x0": 0.018, "x1": 0.185, "y0": 0.112, "y1": 0.160},
        "TITLE_REQUIRED_TOKENS": ["물류센터"],
        "PHOTO_BOX": {"x0": 0.019, "x1": 0.360, "y0": 0.193, "y1": 0.907},
        "SPACE_BOX": {"x0": 0.374, "x1": 0.720, "y0": 0.193, "y1": 0.907},
        "SPACE_PAD": {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0},
        "SPEC_BOX": {"x0": 0.470, "x1": 0.720, "y0": 0.700, "y1": 0.900},
        "SPEC_FIELD": "기타",
        "OCR_LANG": "kor+eng",
    },

    "RSQUARE": {
    "name": "RSQUARE",
    "label": "RSQUARE",

    "keywords": [
        "RSQUARE",
        "물류센터",
        "조감도",
        "물류센터 개요",
        "건축도면",
        "층별정보",
    ],

    # Make sure this exact RSQUARE layout/page is selected
    "PAGE_REQUIRED_TOKENS": [
        "물류센터 개요",
        "위치",
        "대지면적",
        "연면적",
        "사용승인일",
        "층별정보",
    ],
    "PAGE_REQUIRED_MIN": 4,

    # General information table
    # Actual area:
    # x ≈ 344~674 / 1035
    # y ≈ 132~389 / 699
    "GI_BOX": {
        "x0": 0.330,
        "x1": 0.655,
        "y0": 0.185,
        "y1": 0.560,
    },

    # Value column starts around x=432px
    "GI_VALUE_X": 0.417,

    "GI_FIELDS": [
        "위치",
        "대지면적",
        "연면적",
        "센터 규모",
        "사용승인일",
        "총 주차 대수",
        "건폐율",
        "용적률",
        "건축물 용도",
        "주 출입구 방향",
        "층당 화장실 개수",
        "IC",
    ],

    # Title: 이천 대월면 물류센터
    "TITLE_BOX": {
        "x0": 0.020,
        "x1": 0.350,
        "y0": 0.055,
        "y1": 0.135,
    },

    "TITLE_REQUIRED_TOKENS": [
        "물류센터",
    ],

    "TITLE_REQUIRED_TOKENS_OPTIONAL": True,

    "TITLE_REJECT_TOKENS": [
        "Master Of Value Enhancement",
        "Leasing Information",
        "공실 요약",
        "동남권",
    ],

    "TITLE_OCR_PSMS": [7, 6, 11],
    "TITLE_OCR_SCALE": 2,

    # Exterior building image only
    "PHOTO_BOX": {
        "x0": 0.000,
        "x1": 0.320,
        "y0": 0.185,
        "y1": 0.560,
    },

    # Bottom 층별정보 table
    "SPACE_BOX": {
        "x0": 0.330,
        "x1": 0.985,
        "y0": 0.590,
        "y1": 0.975,
    },

    "SPACE_PAD": {
        "left": 0.0,
        "right": 0.0,
        "top": 0.0,
        "bottom": 0.0,
    },

    # 특이사항 column on far right
    "SPEC_BOX": {
        "x0": 0.895,
        "x1": 0.985,
        "y0": 0.595,
        "y1": 0.970,
    },

    "SPEC_FIELD": "특이사항",

    "OCR_LANG": "kor+eng",
},
}

DEFAULT_TEMPLATE = "MOVE"