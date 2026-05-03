"""Keyword-based transaction categoriser with user-override support."""

import json
import re
from pathlib import Path
from typing import Optional

from .parser import Transaction

# Default rules: (category, [keyword patterns])
DEFAULT_RULES: list[tuple[str, list[str]]] = [
    ("Groceries",       ["tesco", "sainsbury", "asda", "waitrose", "lidl", "aldi", "morrisons", "co-op", "coop", "marks & spencer food", "m&s food", "ocado", "whole foods"]),
    ("Eating Out",      ["mcdonalds", "mcdonald", "kfc", "burger king", "subway", "nando", "nandos", "pizza", "deliveroo", "uber eat", "just eat", "costa", "starbucks", "caffe nero", "pret", "greggs", "wagamama", "restaurant", "cafe", "bistro", "diner"]),
    ("Transport",       ["tfl", "transport for london", "national rail", "trainline", "avanti", "gwr", "crosscountry", "uber", "lyft", "bolt", "addison lee", "parking", "petrol", "bp ", "shell ", "esso", "texaco", "fuel"]),
    ("Shopping",        ["amazon", "ebay", "asos", "next ", "h&m", "primark", "zara", "marks & spencer", "m&s", "john lewis", "argos", "currys", "ikea", "tk maxx", "tkmaxx"]),
    ("Entertainment",  ["netflix", "spotify", "apple", "disney", "amazon prime", "sky ", "now tv", "nowtv", "cinema", "odeon", "vue ", "cineworld", "ticketmaster", "eventbrite", "steam", "playstation", "xbox"]),
    ("Health",          ["pharmacy", "chemist", "boots ", "lloyds pharmacy", "superdrug", "nhs", "dentist", "optician", "gym", "pure gym", "puregym", "david lloyd", "virgin active", "leisure centre"]),
    ("Utilities",       ["british gas", "edf", "eon ", "e.on", "octopus energy", "bulb", "severn trent", "thames water", "anglian water", "virgin media", "bt ", "sky broadband", "talk talk", "talktalk", "vodafone", "o2 ", "three ", "ee "]),
    ("Finance",         ["barclays", "hsbc", "lloyds", "natwest", "santander", "monzo", "starling", "revolut", "paypal", "wise", "transfer", "interest", "fee", "charge", "loan", "mortgage", "insurance", "premium", "direct debit"]),
    ("Travel",          ["airbnb", "booking.com", "expedia", "ryanair", "easyjet", "british airways", "ba.com", "hotel", "hostel", "holiday", "flight", "airport"]),
    ("Income",          ["salary", "wages", "payroll", "bacs credit", "bank credit", "dividend", "refund", "cashback", "interest credit"]),
    ("Cash / ATM",      ["cash", "atm", "cashpoint", "withdrawal"]),
    ("Savings",         ["savings", "isa ", "investment", "pension", "vanguard", "hargreaves"]),
]

_CUSTOM_RULES_FILE = Path.home() / ".budget_app" / "categories.json"


def _load_custom_rules() -> list[tuple[str, list[str]]]:
    if not _CUSTOM_RULES_FILE.exists():
        return []
    try:
        data = json.loads(_CUSTOM_RULES_FILE.read_text())
        return [(item["category"], item["keywords"]) for item in data]
    except Exception:
        return []


def save_custom_rule(category: str, keywords: list[str]) -> None:
    _CUSTOM_RULES_FILE.parent.mkdir(parents=True, exist_ok=True)
    existing = _load_custom_rules()
    for i, (cat, kws) in enumerate(existing):
        if cat.lower() == category.lower():
            existing[i] = (category, list(set(kws + keywords)))
            break
    else:
        existing.append((category, keywords))
    data = [{"category": cat, "keywords": kws} for cat, kws in existing]
    _CUSTOM_RULES_FILE.write_text(json.dumps(data, indent=2))


def categorise(transaction: Transaction) -> str:
    desc = transaction.description.lower()
    # Custom rules take priority
    for category, keywords in _load_custom_rules():
        if any(kw.lower() in desc for kw in keywords):
            return category
    # Income heuristic: positive amounts that match income keywords
    for category, keywords in DEFAULT_RULES:
        if category == "Income" and transaction.is_credit:
            if any(kw.lower() in desc for kw in keywords):
                return category
    for category, keywords in DEFAULT_RULES:
        if any(kw.lower() in desc for kw in keywords):
            return category
    # Fallback: credits → Income, debits → Uncategorised
    return "Income" if transaction.is_credit else "Uncategorised"


def categorise_all(transactions: list[Transaction]) -> list[Transaction]:
    for t in transactions:
        if t.category is None:
            t.category = categorise(t)
    return transactions
