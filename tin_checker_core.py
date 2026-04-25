from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Iterable

from pure_python.validators_generated import validate_entry


EN_MESSAGE_REPLACEMENTS = [
    ("Ülke kodu tanınmıyor veya desteklenmiyor.", "Country code is not recognized or supported."),
    ("Ülke kodu zorunludur.", "Country code is required."),
    ("VKN zorunludur.", "TIN is required."),
    ("sadece rakamlardan oluşmalıdır", "must contain digits only"),
    ("tamamen rakam", "digits only"),
    ("tamamen rakamlardan oluşmalıdır", "must contain digits only"),
    ("sadece harf ve rakamlardan oluşmalıdır", "must contain letters and digits only"),
    ("boş olamaz", "cannot be empty"),
    ("geçerli bir tarih olmalıdır", "must be a valid date"),
    ("formatı geçersiz", "format is invalid"),
    ("kontrol basamağı uyuşmuyor", "check digit does not match"),
    ("haneli", "digits"),
    ("karakter", "characters"),
    ("rakam", "digit"),
    ("harf", "letter"),
    ("ile başlamaz", "must not start with"),
    ("ile başlamalı", "must start with"),
    ("uzunluğu geçersiz", "length is invalid"),
    ("geçersiz", "invalid"),
    ("olmalıdır", "is required"),
    ("olmalı", "must be"),
    ("veya", "or"),
    ("için", "for"),
    ("sadece", "only"),
    ("boşluk", "space"),
    ("özel karakter", "special character"),
    ("içeremez", "cannot contain"),
    ("Girilen", "Entered"),
    ("Hesaplanan", "Calculated"),
    ("Desteklenmeyen ülke kodu", "Unsupported country code"),
]

TR_MESSAGE_REPLACEMENTS = [
    ("Country code is required.", "Ülke kodu zorunludur."),
    ("TIN is required.", "TIN zorunludur."),
    ("Validation rule failed safely:", "Doğrulama kuralı güvenli şekilde durduruldu:"),
]

TR_COUNTRY_LABEL_OVERRIDES = {
    "AE": "Birleşik Arap Emirlikleri",
    "CW": "Curaçao",
    "CZ": "Çekya",
    "GB": "Birleşik Krallık",
    "KR": "Güney Kore",
}


def _base_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def _resource_path(relative_path: str) -> Path:
    return _base_dir() / relative_path


def _repair_mojibake(value: str) -> str:
    if not isinstance(value, str):
        return value
    if any(marker in value for marker in ("Ã", "Ä", "Å")):
        try:
            return value.encode("latin1").decode("utf-8")
        except UnicodeError:
            return value
    return value


def _contains_turkish_text(value: str) -> bool:
    lowered = value.lower()
    markers = ("ı", "ğ", "ü", "ş", "ö", "ç", "vkn", "ülke", "hane", "rakam", "geçersiz", "doğru")
    return any(marker in lowered for marker in markers)


def _translate_message(value: str, lang: str, field_name: str) -> str:
    value = _repair_mojibake(value or "")
    if not value:
        return value

    if lang == "tr":
        translated = value
        for source, target in TR_MESSAGE_REPLACEMENTS:
            translated = translated.replace(source, target)
        return translated

    translated = value
    for source, target in EN_MESSAGE_REPLACEMENTS:
        translated = translated.replace(source, target)

    if _contains_turkish_text(translated):
        if field_name == "formatInfo":
            return "Country-specific TIN format rule is applied for the selected country."
        return "The entered TIN does not match the selected country's validation rule."
    return translated


def load_countries(lang: str = "en") -> list[dict[str, str]]:
    path = _resource_path("pure_python/countries_by_lang.json")
    data = json.loads(path.read_text(encoding="utf-8"))
    countries = data.get(lang) or data.get("en") or []
    return [
        {
            "code": country["code"],
            "label": TR_COUNTRY_LABEL_OVERRIDES.get(country["code"], _repair_mojibake(country.get("label", country["code"])))
            if lang == "tr"
            else _repair_mojibake(country.get("label", country["code"])),
        }
        for country in countries
    ]


def country_choices(lang: str = "en") -> list[str]:
    return [f"{country['code']} - {country['label']}" for country in load_countries(lang)]


def extract_country_code(choice: str) -> str:
    return (choice or "").split(" - ", 1)[0].strip().upper()


def _localize_result(result: dict, lang: str) -> dict:
    localized = {}
    for key, value in result.items():
        if isinstance(value, str):
            localized[key] = _translate_message(value, lang, key)
        else:
            localized[key] = value
    return localized


def validate_tin(country_code: str, tin: str, lang: str = "en") -> dict:
    lang = "tr" if lang == "tr" else "en"
    country_code = (country_code or "").strip().upper()
    tin = (tin or "").strip()
    if not country_code:
        return _localize_result({
            "countryCode": "",
            "vkn": tin,
            "isValid": False,
            "formatInfo": "",
            "errorMsg": "Country code is required.",
        }, lang)
    if not tin:
        return _localize_result({
            "countryCode": country_code,
            "vkn": "",
            "isValid": False,
            "formatInfo": "",
            "errorMsg": "TIN is required.",
        }, lang)

    try:
        result = validate_entry(country_code, tin)
    except Exception as exc:
        return _localize_result({
            "countryCode": country_code,
            "vkn": tin,
            "isValid": False,
            "formatInfo": "",
            "errorMsg": f"Validation rule failed safely: {exc}",
        }, lang)
    return _localize_result(result, lang)


def validate_many(country_code: str, tin_values: Iterable[str], lang: str = "en") -> dict:
    results = []
    for tin in tin_values:
        clean_tin = (tin or "").strip()
        if clean_tin:
            results.append(validate_tin(country_code, clean_tin, lang))

    valid_count = sum(1 for item in results if item.get("isValid"))
    return {
        "summary": {
            "total": len(results),
            "valid": valid_count,
            "invalid": len(results) - valid_count,
        },
        "results": results,
    }


def validate_bulk_entries(entries: Iterable[dict], lang: str = "en") -> dict:
    groups = []
    total = 0
    valid = 0

    for entry in entries:
        country_code = (entry.get("countryCode") or "").strip().upper()
        values = entry.get("values") or []
        payload = validate_many(country_code, values, lang)
        if payload["summary"]["total"] == 0 and not country_code:
            continue
        groups.append({"countryCode": country_code, "results": payload["results"]})
        total += payload["summary"]["total"]
        valid += payload["summary"]["valid"]

    return {
        "summary": {"total": total, "valid": valid, "invalid": total - valid},
        "groups": groups,
    }
