"""
eiin_lookup.py - Dynamic Bangladesh Educational Institution EIIN & Address Resolver

Resolves EIIN (Educational Institution Identification Number) and address dynamically
for Bangladesh schools, colleges, and madrasahs using:
1. Explicit EIIN from scraped records, uploaded files, or embedded digits in institute name.
2. Curated institutional registry (major standalone colleges & verified institutions with full addresses).
3. Comprehensive local Bangladesh EIIN database (~29,500+ schools, colleges, and madrasahs).

Important: Never fabricates or falls back to an unrelated institution's EIIN (such as a board default)
when an institution is genuinely not matched. Returns "" for unknown EIINs so callers can display N/A
or allow user input without corrupting data.
"""

from difflib import SequenceMatcher
import json
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_JSON_PATH = os.path.join(BASE_DIR, "data", "bd_institutions_eiin.json")


# Curated registry of verified standalone colleges & notable institutions (with full addresses)
CURATED_INSTITUTIONS: Dict[str, Dict[str, str]] = {
    "132092": {
        "name": "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE",
        "address": "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
        "eiin": "132092",
        "board": "Cumilla",
        "type": "college",
    },
    "105820": {
        "name": "COMILLA AJIT GUHA COLLEGE",
        "address": "Chawkbazar, Cumilla Sadar, Cumilla - 3500, Bangladesh",
        "eiin": "105820",
        "board": "Cumilla",
        "type": "college",
    },
    "105814": {
        "name": "CUMILLA VICTORIA GOVT. COLLEGE",
        "address": "Victoria College Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
        "eiin": "105814",
        "board": "Cumilla",
        "type": "college",
    },
    "105786": {
        "name": "CUMILLA ZILLA SCHOOL",
        "address": "Kandirpar, Cumilla Sadar, Cumilla, Bangladesh",
        "eiin": "105786",
        "board": "Cumilla",
        "type": "school",
    },
    "105767": {
        "name": "COMILLA MODERN HIGH SCHOOL",
        "address": "Nazrul Avenue, Cumilla Sadar, Cumilla - 3500, Bangladesh",
        "eiin": "105767",
        "board": "Cumilla",
        "type": "school",
    },
    "105963": {
        "name": "RAHMANIA CHIRA SABUJ HIGH SCHOOL AND COLLEGE",
        "address": "Lalmai, Cumilla, Bangladesh",
        "eiin": "105963",
        "board": "Cumilla",
        "type": "school_and_college",
    },
    "107109": {
        "name": "DALTA COLLEGE",
        "address": "Dalta, Ramganj, Lakshmipur, Bangladesh",
        "eiin": "107109",
        "board": "Cumilla",
        "type": "college",
    },
    "107660": {
        "name": "BANDERHAT ABDUL MALEQUE COLLEGE",
        "address": "Banderhat, Noakhali Sadar, Noakhali, Bangladesh",
        "eiin": "107660",
        "board": "Cumilla",
        "type": "college",
    },
    "103887": {
        "name": "MATLAB GOVT. DEGREE COLLEGE",
        "address": "Matlab Dakshin, Chandpur, Bangladesh",
        "eiin": "103887",
        "board": "Cumilla",
        "type": "college",
    },
    "103885": {
        "name": "ROYMONEN NESSA MOHILA COLLEGE",
        "address": "Matlab Dakshin, Chandpur, Bangladesh",
        "eiin": "103885",
        "board": "Cumilla",
        "type": "college",
    },
    "107977": {
        "name": "DHAKA COLLEGE",
        "address": "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh",
        "eiin": "107977",
        "board": "Dhaka",
        "type": "college",
    },
    "107975": {
        "name": "DHAKA CITY COLLEGE",
        "address": "Road No 2, Dhanmondi, Dhaka - 1205, Bangladesh",
        "eiin": "107975",
        "board": "Dhaka",
        "type": "college",
    },
    "108274": {
        "name": "NOTRE DAME COLLEGE",
        "address": "Toyenbee Circular Road, Motijheel, Dhaka - 1000, Bangladesh",
        "eiin": "108274",
        "board": "Dhaka",
        "type": "college",
    },
    "108573": {
        "name": "RAJUK UTTARA MODEL COLLEGE",
        "address": "Sector-6, Uttara Model Town, Dhaka - 1230, Bangladesh",
        "eiin": "108573",
        "board": "Dhaka",
        "type": "school_and_college",
    },
    "107845": {
        "name": "ADAMJEE CANTONMENT COLLEGE",
        "address": "Dhaka Cantonment, Dhaka - 1206, Bangladesh",
        "eiin": "107845",
        "board": "Dhaka",
        "type": "college",
    },
    "108348": {
        "name": "VIQARUNNISA NOON COLLEGE",
        "address": "1/A Bailey Road, Ramna, Dhaka - 1000, Bangladesh",
        "eiin": "108348",
        "board": "Dhaka",
        "type": "school_and_college",
    },
    "108272": {
        "name": "HOLY CROSS COLLEGE",
        "address": "Tejgaon, Dhaka - 1215, Bangladesh",
        "eiin": "108272",
        "board": "Dhaka",
        "type": "college",
    },
    "108035": {
        "name": "GOVT. SCIENCE COLLEGE",
        "address": "Farmgate, Tejgaon, Dhaka - 1215, Bangladesh",
        "eiin": "108035",
        "board": "Dhaka",
        "type": "college",
    },
    "132140": {
        "name": "MOHAMMADPUR GOVT. COLLEGE",
        "address": "Satmasjid Road, Mohammadpur, Dhaka - 1207, Bangladesh",
        "eiin": "132140",
        "board": "Dhaka",
        "type": "college",
    },
    "108794": {
        "name": "GOVT. SARADA SUNDARI MOHILA COLLEGE",
        "address": "Faridpur Sadar, Faridpur, Bangladesh",
        "eiin": "108794",
        "board": "Dhaka",
        "type": "college",
    },
    "107859": {
        "name": "B. A. F. SHAHEEN COLLEGE (KURMITOLA)",
        "address": "Dhaka Cantonment, Dhaka - 1206, Bangladesh",
        "eiin": "107859",
        "board": "Dhaka",
        "type": "school_and_college",
    },
    "108163": {
        "name": "AZIMPUR GOVT. GIRLS' SCHOOL AND COLLEGE",
        "address": "Azimpur, Lalbagh, Dhaka - 1205, Bangladesh",
        "eiin": "108163",
        "board": "Dhaka",
        "type": "school_and_college",
    },
    "134599": {
        "name": "HAZI COLLEGE",
        "address": "Dakshinkhan, Dhaka, Bangladesh",
        "eiin": "134599",
        "board": "Dhaka",
        "type": "college",
    },
    "108223": {
        "name": "DHAKA SHIKSHA BOARD LABORATORY SCHOOL & COLLEGE",
        "address": "Mirpur-1, Dhaka - 1216, Bangladesh",
        "eiin": "108223",
        "board": "Dhaka",
        "type": "school_and_college",
    },
    "133008": {
        "name": "CHITTAGONG COLLEGE",
        "address": "College Road, Chawkbazar, Chattogram - 4203, Bangladesh",
        "eiin": "133008",
        "board": "Chattogram",
        "type": "college",
    },
    "138012": {
        "name": "S ALAM COLLEGIATE SCHOOL AND COLLEGE",
        "address": "Karnaphuli, Chattogram, Bangladesh",
        "eiin": "138012",
        "board": "Chattogram",
        "type": "school_and_college",
    },
    "104532": {
        "name": "GOVT. COMMERCE COLLEGE",
        "address": "Agrabad, Chattogram - 4100, Bangladesh",
        "eiin": "104532",
        "board": "Chattogram",
        "type": "college",
    },
    "104533": {
        "name": "GOVT. CITY COLLEGE",
        "address": "Ice Factory Road, Chattogram - 4000, Bangladesh",
        "eiin": "104533",
        "board": "Chattogram",
        "type": "college",
    },
    "104534": {
        "name": "GOVERNMENT MOHSIN COLLEGE",
        "address": "College Road, Chawkbazar, Chattogram - 4203, Bangladesh",
        "eiin": "104534",
        "board": "Chattogram",
        "type": "college",
    },
    "104537": {
        "name": "GOVT. WOMEN'S COLLEGE",
        "address": "Nasirabad, Chattogram, Bangladesh",
        "eiin": "104537",
        "board": "Chattogram",
        "type": "college",
    },
    "104052": {
        "name": "CHITTAGONG CANTONMENT PUBLIC COLLEGE",
        "address": "Chittagong Cantonment, Chattogram - 4220, Bangladesh",
        "eiin": "104052",
        "board": "Chattogram",
        "type": "college",
    },
    "104050": {
        "name": "BAF SHAHEEN COLLEGE CHITTAGONG",
        "address": "Patenga, Chattogram, Bangladesh",
        "eiin": "104050",
        "board": "Chattogram",
        "type": "college",
    },
    "104505": {
        "name": "CHITTAGONG COLLEGIATE SCHOOL",
        "address": "Ice Factory Road, Chattogram, Bangladesh",
        "eiin": "104505",
        "board": "Chattogram",
        "type": "school_and_college",
    },
    "104753": {
        "name": "FAUJDARHAT CADET COLLEGE",
        "address": "Faujdarhat, Sitakunda, Chattogram, Bangladesh",
        "eiin": "104753",
        "board": "Chattogram",
        "type": "college",
    },
    "126490": {
        "name": "RAJSHAHI COLLEGE",
        "address": "Boalia, Rajshahi - 6000, Bangladesh",
        "eiin": "126490",
        "board": "Rajshahi",
        "type": "college",
    },
    "100742": {
        "name": "GOVT. BRAJOMOHUN (BM) COLLEGE",
        "address": "College Road, Barishal - 8200, Bangladesh",
        "eiin": "100742",
        "board": "Barishal",
        "type": "college",
    },
    "130456": {
        "name": "MURARI CHAND (MC) COLLEGE",
        "address": "Tilagarh, Sylhet - 3100, Bangladesh",
        "eiin": "130456",
        "board": "Sylhet",
        "type": "college",
    },
    "127485": {
        "name": "CARMICHAEL COLLEGE",
        "address": "Lalbag, Rangpur - 5400, Bangladesh",
        "eiin": "127485",
        "board": "Rangpur",
        "type": "college",
    },
    "117106": {
        "name": "GOVT. BRAJALAL (BL) COLLEGE",
        "address": "Daulatpur, Khulna - 9202, Bangladesh",
        "eiin": "117106",
        "board": "Khulna",
        "type": "college",
    },
    "111831": {
        "name": "GOVT. ANANDA MOHAN COLLEGE",
        "address": "College Road, Mymensingh - 2200, Bangladesh",
        "eiin": "111831",
        "board": "Mymensingh",
        "type": "college",
    },
    "106579": {
        "name": "FENI GOVT. COLLEGE",
        "address": "College Road, Feni - 3900, Bangladesh",
        "eiin": "106579",
        "board": "Cumilla",
        "type": "college",
    },
    "103284": {
        "name": "BRAHMANBARIA GOVT. COLLEGE",
        "address": "College Road, Brahmanbaria - 3400, Bangladesh",
        "eiin": "103284",
        "board": "Cumilla",
        "type": "college",
    },
    "103597": {
        "name": "CHANDPUR GOVT. COLLEGE",
        "address": "Stadium Road, Chandpur - 3600, Bangladesh",
        "eiin": "103597",
        "board": "Cumilla",
        "type": "college",
    },
    "107541": {
        "name": "NOAKHALI GOVT. COLLEGE",
        "address": "Maijdee Court, Noakhali - 3800, Bangladesh",
        "eiin": "107541",
        "board": "Cumilla",
        "type": "college",
    },
    "107386": {
        "name": "BANDERHAT HIGH SCHOOL & COLLEGE",
        "address": "Banderhat, Lakshmipur, Bangladesh",
        "eiin": "107386",
        "board": "Cumilla",
        "type": "school_and_college",
    },
    "115957": {
        "name": "GOVT. M. M. COLLEGE, JASHORE",
        "address": "Kharki, Jashore - 7400, Bangladesh",
        "eiin": "115957",
        "board": "Jessore",
        "type": "college",
    },
    "120583": {
        "name": "DINAJPUR GOVT. COLLEGE",
        "address": "Sadar, Dinajpur - 5200, Bangladesh",
        "eiin": "120583",
        "board": "Dinajpur",
        "type": "college",
    },
    "108156": {
        "name": "GOVT. MADRASA-E-ALIA DHAKA",
        "address": "Bakshibazar, Dhaka - 1211, Bangladesh",
        "eiin": "108156",
        "board": "Madrasah",
        "type": "madrasah",
    },
    "108036": {
        "name": "DHAKA POLYTECHNIC INSTITUTE",
        "address": "Tejgaon I/A, Dhaka - 1208, Bangladesh",
        "eiin": "108036",
        "board": "Technical",
        "type": "college",
    },
}

# Explicit aliases for common board name variations -> canonical EIIN
CURATED_ALIASES: Dict[str, str] = {
    "CUMILLA GOVT MODEL COLLEGE": "132092",
    "COMILLA GOVT MODEL COLLEGE": "132092",
    "CUMILLA SHIKKHA BOARD MODEL COLLEGE": "132092",
    "COMILLA SHIKKHA BOARD MODEL COLLEGE": "132092",
    "COMILLA SHIKKHA BOARD GOVT MODEL COLLEGE": "132092",
    "BM COLLEGE": "100742",
    "S ALAM COLLEGIATE SCHOOL AND COLLEGE": "138012",
    "S ALAM COLLEGIATE SCHOOL AND COLLEGE CHATTOGRAM": "138012",
    "S. ALAM COLLEGIATE SCHOOL & COLLEGE": "138012",
    "CHITTAGONG COLLEGE": "133008",
    "CHATTOGRAM COLLEGE": "133008",
    "GOVT COMMERCE COLLEGE": "104532",
    "GOVT COMMERCE COLLEGE CHITTAGONG": "104532",
    "GOVT CITY COLLEGE": "104533",
    "GOVT CITY COLLEGE CHATTOGRAM": "104533",
    "GOVT MOHSIN COLLEGE": "104534",
    "GOVT HAZI MUHAMMAD MOHSIN COLLEGE": "104534",
    "CHITTAGONG CANTONMENT PUBLIC COLLEGE": "104052",
    "FAUJDARHAT CADET COLLEGE": "104753",
    "MC COLLEGE": "130456",
    "BL COLLEGE": "117106",
    "MM COLLEGE": "115957",
    "GOVT MM COLLEGE": "115957",
    "GOVT BM COLLEGE": "100742",
    "GOVT BL COLLEGE": "117106",
    "BANDERHAT ABDUL MALEQUE COLLEGE": "107660",
    "DALTA DEGREE COLLEGE": "107109",
    "DALTA COLLEGE": "107109",
    "BAF SHAHEEN COLLEGE KURMITOLA": "107859",
    "GOVT AZIMPUR GIRLS SCHOOL AND COLLEGE": "108163",
    "AZIMPUR GOVT GIRLS SCHOOL AND COLLEGE": "108163",
}


BD_DISTRICTS: List[str] = [
    "Bagerhat", "Bandarban", "Barguna", "Barishal", "Barisal", "Bhola", "Bogra", "Bogura", "Brahmanbaria",
    "Chandpur", "Chattogram", "Chittagong", "Chuadanga", "Comilla", "Cumilla", "Cox's Bazar",
    "Dhaka", "Dinajpur", "Faridpur", "Feni", "Gaibandha", "Gazipur", "Gopalganj",
    "Habiganj", "Jamalpur", "Jashore", "Jessore", "Jhalokati", "Jhenaidah", "Joypurhat",
    "Khagrachhari", "Khulna", "Kishoreganj", "Kurigram", "Kushtia", "Lakshmipur", "Laxmipur",
    "Lalmonirhat", "Madaripur", "Magura", "Manikganj", "Meherpur", "Moulvibazar", "Munshiganj",
    "Mymensingh", "Naogaon", "Narail", "Narayanganj", "Narsingdi", "Natore", "Netrokona",
    "Nilphamari", "Noakhali", "Pabna", "Panchagarh", "Patuakhali", "Pirojpur", "Rajbari",
    "Rajshahi", "Rangamati", "Rangpur", "Satkhira", "Shariatpur", "Sherpur", "Sirajganj",
    "Sunamganj", "Sylhet", "Tangail", "Thakurgaon",
]

DISTRICT_CANONICAL: Dict[str, str] = {
    "COMILLA": "CUMILLA",
    "CHITTAGONG": "CHATTOGRAM",
    "JESSORE": "JASHORE",
    "BARISAL": "BARISHAL",
    "BOGRA": "BOGURA",
    "LAXMIPUR": "LAKSHMIPUR",
}

DISTRICT_DISPLAY: Dict[str, str] = {
    "CUMILLA": "Cumilla",
    "CHATTOGRAM": "Chattogram",
    "JASHORE": "Jashore",
    "BARISHAL": "Barishal",
    "BOGURA": "Bogura",
    "LAKSHMIPUR": "Lakshmipur",
}

BD_DISTRICTS_UPPER: Set[str] = {d.upper() for d in BD_DISTRICTS} | set(DISTRICT_CANONICAL.values())

GENERIC_TOKENS: Set[str] = {
    "HIGH", "SCHOOL", "COLLEGE", "GOVERNMENT", "GOVT", "AND", "SECONDARY",
    "JUNIOR", "PRIMARY", "MODEL", "DEGREE", "UNIVERSITY", "INSTITUTE", "INSTITUTION",
    "ACADEMY", "MADRASAH", "DAKHIL", "ALIM", "FAZIL", "KAMIL", "GIRLS", "BOYS",
    "WOMEN", "WOMENS", "MAHILA", "PUBLIC", "PILOT", "ADARSHA", "IDEAL",
    "MEMORIAL", "TECHNICAL", "POLYTECHNIC", "COLLEGIATE", "LABORATORY",
    "CANTONMENT", "BOARD", "EDUCATION", "SHIKKHA", "SHIKSHA", "THE", "OF",
    "SADAR", "POURASHAVA", "MUNICIPAL", "UNION", "BAZAR", "HAT", "GANJ",
} | BD_DISTRICTS_UPPER

SPELLING_REPLACEMENTS: Dict[str, str] = {
    "GOVT": "GOVERNMENT",
    "COLL": "COLLEGE",
    "CLG": "COLLEGE",
    "SCH": "SCHOOL",
    "INST": "INSTITUTE",
    "MAD": "MADRASAH",
    "MADRASA": "MADRASAH",
    "MADRASHA": "MADRASAH",
    "MADRASHAH": "MADRASAH",
    "COMILLA": "CUMILLA",
    "CHITTAGONG": "CHATTOGRAM",
    "JESSORE": "JASHORE",
    "BARISAL": "BARISHAL",
    "BOGRA": "BOGURA",
    "LAXMIPUR": "LAKSHMIPUR",
    "SABUZ": "SABUJ",
    "SHABUZ": "SABUJ",
    "SHABUJ": "SABUJ",
    "MOHILA": "MAHILA",
    "CHOWDHURY": "CHOUDHURY",
    "CHOWDHURI": "CHOUDHURY",
    "MOJUMDER": "MAZUMDER",
    "MAJUMDER": "MAZUMDER",
    "MOZUMDER": "MAZUMDER",
    "HOSSEN": "HOSSAIN",
    "HUSAIN": "HOSSAIN",
    "HUSSAIN": "HOSSAIN",
    "HOSEN": "HOSSAIN",
    "ROHMAN": "RAHMAN",
    "ISLAMIYA": "ISLAMIA",
    "IDEL": "IDEAL",
    "SHIKSHA": "SHIKKHA",
    "SHIKHA": "SHIKKHA",
    "SIKKHA": "SHIKKHA",
    "ZILA": "ZILLA",
    "JILA": "ZILLA",
    "JELA": "ZILLA",
    "MALEK": "MALEQUE",
    "MALIK": "MALEQUE",
    "KHALEK": "KHALEQUE",
    "SADEK": "SADEQUE",
    "ALAM": "ALAM",
    "GIRL": "GIRLS",
    "BOY": "BOYS",
    "WOMENS": "WOMEN",
    "ACADEMI": "ACADEMY",
    "ACADAMY": "ACADEMY",
    "COLLEG": "COLLEGE",
    "SCOL": "SCHOOL",
    "SECONDERY": "SECONDARY",
    "SECONDRY": "SECONDARY",
}


def normalize_institute_name(raw_name: str) -> Tuple[str, str]:
    """
    Normalizes an institution name for reliable cross-source matching.
    Returns (normalized_base_name, extracted_district_upper).
    """
    s = str(raw_name or "").upper().strip()
    if not s or s in {"N/A", "NAN", "NONE", "-"}:
        return "", ""

    # Strip embedded EIIN digits (5-7 digits)
    s = re.sub(r"\(?\b\d{5,7}\b\)?", " ", s)

    # Extract trailing district after comma, e.g. "..., COMILLA" or "..., LAXMIPUR"
    extracted_district = ""
    if "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if len(parts) >= 2:
            tail = re.sub(r"[^A-Z\s']", "", parts[-1]).strip()
            tail_canon = DISTRICT_CANONICAL.get(tail, tail)
            if tail_canon in BD_DISTRICTS_UPPER:
                extracted_district = tail_canon
                s = " ".join(parts[:-1])

    # Expand common board abbreviations before stripping punctuation
    # e.g. "B. A. F." -> "BAF", "M. M." -> "MM"
    s = re.sub(r"\bB\s*\.\s*A\s*\.\s*F\s*\.?", " BAF ", s)
    s = re.sub(r"\bH\s*/\s*S\b", " HIGH SCHOOL ", s)
    s = re.sub(r"\bH\s*\.\s*S\s*\.", " HIGH SCHOOL ", s)
    s = re.sub(r"\bH\s+S\b", " HIGH SCHOOL ", s)
    s = re.sub(r"\bP\s*\.\s*S\s*\.", " PILOT SCHOOL ", s)
    s = re.sub(r"\bG\s*\.\s*H\s*\.\s*S\b", " GOVERNMENT HIGH SCHOOL ", s)
    s = re.sub(r"&", " AND ", s)

    # Replace remaining punctuation with spaces
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    raw_tokens = [t for t in s.split() if t]

    # Map spelling variants and detect district mentions anywhere
    norm_tokens: List[str] = []
    for tok in raw_tokens:
        mapped = SPELLING_REPLACEMENTS.get(tok, tok)
        if not extracted_district and mapped in BD_DISTRICTS_UPPER:
            extracted_district = mapped
        norm_tokens.append(mapped)

    return " ".join(norm_tokens), extracted_district


def _classify_inst_type(norm_name: str, declared_type: str = "") -> str:
    has_school = "SCHOOL" in norm_name
    has_college = "COLLEGE" in norm_name
    has_madrasah = any(k in norm_name for k in ("MADRASAH", "DAKHIL", "ALIM", "FAZIL", "KAMIL"))
    if has_madrasah:
        return "madrasah"
    if has_school and has_college:
        return "school_and_college"
    if has_college:
        return "college"
    if has_school:
        return "school"
    dt = (declared_type or "").lower()
    if "mad" in dt:
        return "madrasah"
    if "college" in dt:
        return "college"
    if "school" in dt:
        return "school"
    return "other"


def _types_compatible(q_type: str, c_type: str) -> bool:
    if q_type == "other" or c_type == "other":
        return True
    if q_type == c_type:
        return True
    # Allow school_and_college to match college (many BD datasets tag School & College as "college")
    if {q_type, c_type} == {"school_and_college", "college"}:
        return True
    if {q_type, c_type} == {"school_and_college", "school"}:
        return True
    return False


class _InstitutionDatabase:
    """Lazy-loaded in-memory index of ~30,000 Bangladesh educational institutions."""

    def __init__(self) -> None:
        self._loaded = False
        self.by_eiin: Dict[str, Dict[str, str]] = {}
        self.exact_index: Dict[str, List[str]] = {}
        self.token_index: Dict[str, List[str]] = {}
        self.norm_meta: Dict[str, Tuple[str, str, str, Set[str]]] = {}  # eiin -> (norm, dist, type, distinctive_tokens)
        self._resolve_cache: Dict[Tuple[str, str, str], Dict[str, str]] = {}

    def _ensure_loaded(self) -> None:
        if self._loaded:
            return
        self._loaded = True

        # 1. Load comprehensive local JSON dataset (~29,500+ entries)
        if os.path.exists(DB_JSON_PATH):
            try:
                with open(DB_JSON_PATH, "r", encoding="utf-8") as f:
                    raw_db = json.load(f)
                for eiin, vals in raw_db.items():
                    if not isinstance(vals, list) or not vals:
                        continue
                    name = str(vals[0]).strip()
                    itype = str(vals[1]).strip() if len(vals) > 1 else ""
                    if eiin and name:
                        self.by_eiin[eiin] = {
                            "eiin": eiin,
                            "name": name,
                            "type": itype,
                            "address": "",
                            "board": "",
                        }
            except Exception:
                pass

        # 2. Merge curated institutions (overrides/adds full address & standalone colleges)
        for eiin, rec in CURATED_INSTITUTIONS.items():
            existing = self.by_eiin.get(eiin, {})
            self.by_eiin[eiin] = {
                "eiin": eiin,
                "name": rec["name"],
                "type": rec.get("type") or existing.get("type", ""),
                "address": rec.get("address", ""),
                "board": rec.get("board", ""),
            }

        # 3. Build fast lookup indices
        for eiin, rec in self.by_eiin.items():
            norm, dist = normalize_institute_name(rec["name"])
            if not norm:
                continue
            itype = _classify_inst_type(norm, rec.get("type", ""))
            tokens = set(norm.split())
            distinctive = {t for t in tokens if len(t) >= 3 and t not in GENERIC_TOKENS}
            if not distinctive:
                distinctive = {t for t in tokens if len(t) >= 2}

            self.norm_meta[eiin] = (norm, dist, itype, distinctive)
            self.exact_index.setdefault(norm, []).append(eiin)
            for tok in distinctive:
                self.token_index.setdefault(tok, []).append(eiin)

        # 4. Register explicit aliases
        for alias, eiin in CURATED_ALIASES.items():
            norm_alias, _ = normalize_institute_name(alias)
            if norm_alias and eiin in self.by_eiin:
                lst = self.exact_index.setdefault(norm_alias, [])
                if eiin not in lst:
                    lst.insert(0, eiin)

    def lookup_by_eiin(self, eiin: str) -> Optional[Dict[str, str]]:
        self._ensure_loaded()
        clean = str(eiin or "").strip()
        if not clean:
            return None
        rec = self.by_eiin.get(clean)
        if not rec:
            return None
        result = dict(rec)
        if not result.get("address"):
            result["address"] = derive_address_from_name(result.get("name", ""), result.get("board", ""))
        return result

    def match_by_name(self, institute: str, board: str = "") -> Optional[Dict[str, str]]:
        self._ensure_loaded()
        q_norm, q_dist = normalize_institute_name(institute)
        if not q_norm:
            return None

        q_type = _classify_inst_type(q_norm)
        board_canon = DISTRICT_CANONICAL.get(str(board or "").upper().strip(), str(board or "").upper().strip())

        def _pick_best(eiin_list: List[str]) -> Optional[str]:
            if not eiin_list:
                return None
            if len(eiin_list) == 1:
                return eiin_list[0]
            # Prefer curated or district/board match
            scored = []
            for e in eiin_list:
                _, c_dist, _, _ = self.norm_meta.get(e, ("", "", "other", set()))
                rec = self.by_eiin[e]
                sc = 0
                if e in CURATED_INSTITUTIONS:
                    sc += 3
                if q_dist and (c_dist == q_dist or q_dist in rec.get("address", "").upper()):
                    sc += 4
                if board_canon and board_canon in (rec.get("board", "").upper() + " " + rec.get("address", "").upper()):
                    sc += 2
                scored.append((sc, e))
            scored.sort(key=lambda x: x[0], reverse=True)
            # Only return if clearly disambiguated or single top curated entry
            if scored[0][0] > scored[1][0]:
                return scored[0][1]
            # If tied and both exist, avoid returning a random wrong district's school
            return None

        # Step A: Exact normalized match
        if q_norm in self.exact_index:
            picked = _pick_best(self.exact_index[q_norm])
            if picked:
                return self.lookup_by_eiin(picked)

        # Also try removing trailing district token if embedded inside q_norm without comma
        q_tokens_list = q_norm.split()
        if len(q_tokens_list) >= 3 and q_tokens_list[-1] in BD_DISTRICTS_UPPER:
            trimmed = " ".join(q_tokens_list[:-1])
            if trimmed in self.exact_index:
                picked = _pick_best(self.exact_index[trimmed])
                if picked:
                    return self.lookup_by_eiin(picked)

        # Step B: Token-indexed fuzzy matching
        q_tokens = set(q_tokens_list)
        q_distinctive = {t for t in q_tokens if len(t) >= 3 and t not in GENERIC_TOKENS}
        if not q_distinctive:
            q_distinctive = {t for t in q_tokens if len(t) >= 2}

        candidate_hits: Dict[str, int] = {}
        for tok in q_distinctive:
            for e in self.token_index.get(tok, []):
                candidate_hits[e] = candidate_hits.get(e, 0) + 1

        if not candidate_hits:
            return None

        # Rank top candidates by distinctive overlap then SequenceMatcher
        min_hits = 1 if len(q_distinctive) <= 2 else max(2, len(q_distinctive) - 1)
        filtered = [e for e, cnt in candidate_hits.items() if cnt >= min_hits]
        if not filtered:
            filtered = [e for e, cnt in sorted(candidate_hits.items(), key=lambda x: -x[1])[:40]]
        elif len(filtered) > 120:
            filtered = [e for e, _ in sorted(((e, candidate_hits[e]) for e in filtered), key=lambda x: -x[1])[:120]]

        best_score = 0.0
        best_eiin: Optional[str] = None
        second_score = 0.0

        for e in filtered:
            c_norm, c_dist, c_type, c_distinctive = self.norm_meta[e]
            if not _types_compatible(q_type, c_type):
                continue

            # Distinctive token Jaccard & coverage
            inter = len(q_distinctive & c_distinctive)
            union = len(q_distinctive | c_distinctive)
            if union == 0 or inter == 0:
                continue

            dist_jaccard = inter / union
            dist_recall = inter / len(q_distinctive)
            seq_ratio = SequenceMatcher(None, q_norm, c_norm).ratio()

            # Weighted score favoring distinctive name match + sequence similarity
            score = (0.45 * seq_ratio) + (0.35 * dist_jaccard) + (0.20 * dist_recall)

            # Small bonus for matching district or curated registry
            rec = self.by_eiin[e]
            if q_dist and (c_dist == q_dist or q_dist in rec.get("address", "").upper() or q_dist in c_norm):
                score += 0.04
            elif q_dist and c_dist and c_dist != q_dist:
                # Conflicting explicit districts -> strong penalty
                score -= 0.25

            if score > best_score:
                second_score = best_score
                best_score = score
                best_eiin = e
            elif score > second_score:
                second_score = score

        # Strict acceptance threshold to prevent false EIIN assignments
        if best_eiin and best_score >= 0.84:
            # Ensure we aren't guessing blindly between two nearly identical ambiguous schools
            if (best_score - second_score) < 0.015 and best_score < 0.95:
                return None
            return self.lookup_by_eiin(best_eiin)

        return None

    def resolve(self, raw_eiin: str = "", institute: str = "", board: str = "") -> Dict[str, str]:
        self._ensure_loaded()
        cache_key = (str(raw_eiin or "").strip(), str(institute or "").strip(), str(board or "").strip())
        if cache_key in self._resolve_cache:
            return dict(self._resolve_cache[cache_key])

        val = cache_key[0]
        inst = cache_key[1]
        brd = cache_key[2]

        clean_eiin = val if (val.isdigit() and 4 <= len(val) <= 8) else ""

        # Check embedded 5-7 digit EIIN inside institute name, e.g. "(105963) ..."
        clean_inst_name = inst
        if inst:
            m = re.search(r"\(?\b(\d{5,7})\b\)?", inst)
            if m:
                if not clean_eiin:
                    clean_eiin = m.group(1)
                clean_inst_name = re.sub(r"\(?\b\d{5,7}\b\)?", "", inst)
                clean_inst_name = re.sub(r"[-()]+", " ", clean_inst_name)
                clean_inst_name = re.sub(r"\s+", " ", clean_inst_name).strip()

        matched: Optional[Dict[str, str]] = None
        if clean_eiin:
            matched = self.lookup_by_eiin(clean_eiin)
        elif clean_inst_name:
            matched = self.match_by_name(clean_inst_name, brd)
            if matched:
                clean_eiin = matched["eiin"]

        final_name = clean_inst_name or (matched["name"] if matched else "")
        final_address = ""
        if matched and matched.get("address") and matched.get("address") != "Bangladesh":
            final_address = matched["address"]
        else:
            final_address = derive_address_from_name(final_name or inst, brd)

        final_board = brd or (matched.get("board", "") if matched else "")

        result = {
            "eiin": clean_eiin,
            "name": final_name,
            "address": final_address,
            "board": final_board,
            "matched": "true" if matched else "false",
        }
        if len(self._resolve_cache) < 2000:
            self._resolve_cache[cache_key] = result
        return dict(result)


def derive_address_from_name(institute: str = "", board: str = "") -> str:
    """Derives a clean geographic address from district mentions in institute name or board."""
    inst_str = str(institute or "").strip()
    if inst_str:
        for dist in BD_DISTRICTS:
            if re.search(rf"\b{re.escape(dist)}\b", inst_str, re.IGNORECASE):
                canon = DISTRICT_CANONICAL.get(dist.upper(), dist.upper())
                disp = DISTRICT_DISPLAY.get(canon, dist)
                return f"{disp}, Bangladesh"

    b_clean = re.sub(r"\s*\(.*?\)", "", str(board or "")).strip()
    b_lower = b_clean.lower()
    board_locations = {
        "cumilla": "Cumilla, Bangladesh",
        "comilla": "Cumilla, Bangladesh",
        "dhaka": "Dhaka, Bangladesh",
        "chattogram": "Chattogram, Bangladesh",
        "chittagong": "Chattogram, Bangladesh",
        "rajshahi": "Rajshahi, Bangladesh",
        "barishal": "Barishal, Bangladesh",
        "barisal": "Barishal, Bangladesh",
        "sylhet": "Sylhet, Bangladesh",
        "dinajpur": "Dinajpur, Bangladesh",
        "jessore": "Jashore, Bangladesh",
        "jashore": "Jashore, Bangladesh",
        "mymensingh": "Mymensingh, Bangladesh",
        "madrasah": "Dhaka, Bangladesh",
        "technical": "Dhaka, Bangladesh",
    }
    for k, loc in board_locations.items():
        if k in b_lower:
            return loc

    return "Bangladesh"


_DB = _InstitutionDatabase()


def lookup_institution_by_eiin(eiin: str) -> Optional[Dict[str, str]]:
    return _DB.lookup_by_eiin(eiin)


def resolve_institution_details(raw_eiin: str = "", institute: str = "", board: str = "") -> Dict[str, str]:
    return _DB.resolve(raw_eiin=raw_eiin, institute=institute, board=board)


def resolve_institution_eiin_dynamic(raw_val: str = "", institute: str = "", board: str = "") -> str:
    return _DB.resolve(raw_eiin=raw_val, institute=institute, board=board)["eiin"]


def resolve_institution_address_dynamic(institute: str = "", eiin: str = "", board: str = "") -> str:
    return _DB.resolve(raw_eiin=eiin, institute=institute, board=board)["address"]
