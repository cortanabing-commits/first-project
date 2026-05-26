"""Expenditure analysis: summaries, trends, and insights."""

from collections import defaultdict
from datetime import datetime
from typing import Optional

from .parser import Transaction


def spending_by_category(transactions: list[Transaction]) -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)
    for t in transactions:
        if t.is_debit:
            totals[t.category or "Uncategorised"] += abs(t.amount)
    return dict(sorted(totals.items(), key=lambda x: x[1], reverse=True))


def income_total(transactions: list[Transaction]) -> float:
    return sum(t.amount for t in transactions if t.is_credit)


def expenditure_total(transactions: list[Transaction]) -> float:
    return sum(abs(t.amount) for t in transactions if t.is_debit)


def net_position(transactions: list[Transaction]) -> float:
    return sum(t.amount for t in transactions)


def monthly_summary(transactions: list[Transaction]) -> dict[str, dict]:
    """Return per-month income, expenditure, and net."""
    months: dict[str, dict] = defaultdict(lambda: {"income": 0.0, "expenditure": 0.0, "net": 0.0})
    for t in transactions:
        key = t.date.strftime("%Y-%m")
        if t.is_credit:
            months[key]["income"] += t.amount
        else:
            months[key]["expenditure"] += abs(t.amount)
        months[key]["net"] += t.amount
    return dict(sorted(months.items()))


def monthly_category_breakdown(transactions: list[Transaction]) -> dict[str, dict[str, float]]:
    """Return {month: {category: amount}} for debits only."""
    result: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for t in transactions:
        if t.is_debit:
            month = t.date.strftime("%Y-%m")
            result[month][t.category or "Uncategorised"] += abs(t.amount)
    return {k: dict(v) for k, v in sorted(result.items())}


def top_merchants(transactions: list[Transaction], n: int = 10) -> list[tuple[str, float, int]]:
    """Return top n merchants by total spend: (description, total, count)."""
    merchants: dict[str, list[float]] = defaultdict(list)
    for t in transactions:
        if t.is_debit:
            merchants[t.description].append(abs(t.amount))
    ranked = [(desc, sum(amts), len(amts)) for desc, amts in merchants.items()]
    ranked.sort(key=lambda x: x[1], reverse=True)
    return ranked[:n]


def largest_transactions(transactions: list[Transaction], n: int = 10) -> list[Transaction]:
    debits = [t for t in transactions if t.is_debit]
    return sorted(debits, key=lambda t: abs(t.amount), reverse=True)[:n]


def date_range(transactions: list[Transaction]) -> tuple[Optional[datetime], Optional[datetime]]:
    if not transactions:
        return None, None
    dates = [t.date for t in transactions]
    return min(dates), max(dates)


def savings_rate(transactions: list[Transaction]) -> Optional[float]:
    inc = income_total(transactions)
    if inc == 0:
        return None
    return (net_position(transactions) / inc) * 100


def generate_insights(transactions: list[Transaction]) -> list[str]:
    insights: list[str] = []
    inc = income_total(transactions)
    exp = expenditure_total(transactions)
    net = net_position(transactions)
    by_cat = spending_by_category(transactions)

    if inc > 0:
        rate = (net / inc) * 100
        if rate >= 20:
            insights.append(f"Great savings rate of {rate:.1f}% — you're saving more than 20% of income.")
        elif rate >= 0:
            insights.append(f"Savings rate: {rate:.1f}%. Aim for 20%+ for a healthy financial cushion.")
        else:
            insights.append(f"You spent {abs(rate):.1f}% more than you earned this period — watch out for overspend.")

    if by_cat:
        top_cat, top_amt = list(by_cat.items())[0]
        if inc > 0:
            pct = (top_amt / inc) * 100
            insights.append(f"Largest spending category is '{top_cat}' at £{top_amt:,.2f} ({pct:.1f}% of income).")

    months = monthly_summary(transactions)
    if len(months) >= 2:
        values = list(months.values())
        trend = values[-1]["expenditure"] - values[-2]["expenditure"]
        direction = "up" if trend > 0 else "down"
        insights.append(f"Monthly spending is {direction} £{abs(trend):,.2f} vs last month.")

    uncategorised = by_cat.get("Uncategorised", 0)
    if uncategorised > 0:
        insights.append(f"£{uncategorised:,.2f} is uncategorised — use 'budget add-rule' to improve accuracy.")

    return insights
