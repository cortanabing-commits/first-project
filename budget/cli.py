"""CLI entry point for the personal budget app."""

from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich import box
from rich.text import Text

from .parser import parse_csv
from .categoriser import categorise_all, save_custom_rule
from .analysis import (
    spending_by_category,
    income_total,
    expenditure_total,
    net_position,
    monthly_summary,
    monthly_category_breakdown,
    top_merchants,
    largest_transactions,
    date_range,
    savings_rate,
    generate_insights,
)

app = typer.Typer(help="Personal budget analyser — ingest bank statements and understand your spending.", no_args_is_help=True)
console = Console()

CURRENCY = "£"


def _fmt(amount: float) -> str:
    return f"{CURRENCY}{amount:,.2f}"


def _load(files: list[Path]) -> list:
    all_txns = []
    for f in files:
        try:
            txns = parse_csv(f)
            categorise_all(txns)
            all_txns.extend(txns)
            console.print(f"[green]✓[/green] Loaded [bold]{len(txns)}[/bold] transactions from [cyan]{f.name}[/cyan]")
        except Exception as e:
            console.print(f"[red]✗[/red] Failed to load {f}: {e}")
    all_txns.sort(key=lambda t: t.date)
    return all_txns


@app.command()
def summary(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    month: Optional[str] = typer.Option(None, "--month", "-m", help="Filter to a specific month (YYYY-MM)"),
):
    """Show an overall financial summary for the statement(s)."""
    txns = _load(files)
    if not txns:
        console.print("[red]No transactions loaded.[/red]")
        raise typer.Exit(1)

    if month:
        txns = [t for t in txns if t.date.strftime("%Y-%m") == month]
        if not txns:
            console.print(f"[yellow]No transactions found for {month}.[/yellow]")
            raise typer.Exit()

    start, end = date_range(txns)
    inc = income_total(txns)
    exp = expenditure_total(txns)
    net = net_position(txns)
    rate = savings_rate(txns)

    console.print()
    period = f"{start.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}" if start and end else "Unknown"
    console.print(Panel(f"[bold]Period:[/bold] {period}   [bold]Transactions:[/bold] {len(txns)}", title="Budget Summary", style="blue"))

    table = Table(box=box.ROUNDED, show_header=False, padding=(0, 2))
    table.add_column(style="bold")
    table.add_column(justify="right")
    table.add_row("Total Income", f"[green]{_fmt(inc)}[/green]")
    table.add_row("Total Expenditure", f"[red]{_fmt(exp)}[/red]")
    net_colour = "green" if net >= 0 else "red"
    table.add_row("Net Position", f"[{net_colour}]{_fmt(net)}[/{net_colour}]")
    if rate is not None:
        table.add_row("Savings Rate", f"[{'green' if rate >= 0 else 'red'}]{rate:.1f}%[/{'green' if rate >= 0 else 'red'}]")
    console.print(table)

    # Insights
    insights = generate_insights(txns)
    if insights:
        console.print()
        console.print(Panel("\n".join(f"  • {i}" for i in insights), title="Insights", style="yellow"))

    console.print()


@app.command()
def categories(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    month: Optional[str] = typer.Option(None, "--month", "-m", help="Filter to a specific month (YYYY-MM)"),
    top: int = typer.Option(0, "--top", "-n", help="Show only top N categories"),
):
    """Break down spending by category."""
    txns = _load(files)
    if not txns:
        raise typer.Exit(1)

    if month:
        txns = [t for t in txns if t.date.strftime("%Y-%m") == month]

    by_cat = spending_by_category(txns)
    total_exp = sum(by_cat.values())

    if top:
        items = list(by_cat.items())[:top]
    else:
        items = list(by_cat.items())

    table = Table(title="Spending by Category", box=box.ROUNDED)
    table.add_column("Category", style="bold")
    table.add_column("Amount", justify="right")
    table.add_column("% of Spend", justify="right")
    table.add_column("Bar")

    for cat, amt in items:
        pct = (amt / total_exp * 100) if total_exp else 0
        bar_len = int(pct / 2)
        bar = "█" * bar_len
        table.add_row(cat, f"[red]{_fmt(amt)}[/red]", f"{pct:.1f}%", f"[cyan]{bar}[/cyan]")

    console.print()
    console.print(table)
    console.print(f"\n  [bold]Total Expenditure:[/bold] [red]{_fmt(total_exp)}[/red]")
    console.print()


@app.command()
def monthly(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    breakdown: bool = typer.Option(False, "--breakdown", "-b", help="Show per-category breakdown per month"),
):
    """Show month-by-month income and expenditure."""
    txns = _load(files)
    if not txns:
        raise typer.Exit(1)

    months = monthly_summary(txns)

    table = Table(title="Monthly Summary", box=box.ROUNDED)
    table.add_column("Month", style="bold")
    table.add_column("Income", justify="right")
    table.add_column("Expenditure", justify="right")
    table.add_column("Net", justify="right")

    for m, data in months.items():
        net = data["net"]
        net_str = f"[{'green' if net >= 0 else 'red'}]{_fmt(net)}[/{'green' if net >= 0 else 'red'}]"
        table.add_row(m, f"[green]{_fmt(data['income'])}[/green]", f"[red]{_fmt(data['expenditure'])}[/red]", net_str)

    console.print()
    console.print(table)

    if breakdown:
        mbd = monthly_category_breakdown(txns)
        for m, cats in mbd.items():
            cat_table = Table(title=f"{m} — Category Breakdown", box=box.SIMPLE)
            cat_table.add_column("Category")
            cat_table.add_column("Amount", justify="right")
            for cat, amt in sorted(cats.items(), key=lambda x: x[1], reverse=True):
                cat_table.add_row(cat, f"[red]{_fmt(amt)}[/red]")
            console.print(cat_table)

    console.print()


@app.command()
def merchants(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    top: int = typer.Option(10, "--top", "-n", help="Number of top merchants to show"),
):
    """List the merchants you spend the most at."""
    txns = _load(files)
    if not txns:
        raise typer.Exit(1)

    ranked = top_merchants(txns, n=top)

    table = Table(title=f"Top {top} Merchants by Spend", box=box.ROUNDED)
    table.add_column("#", justify="right", style="dim")
    table.add_column("Merchant")
    table.add_column("Total Spent", justify="right")
    table.add_column("Transactions", justify="right")
    table.add_column("Avg per Visit", justify="right")

    for i, (desc, total, count) in enumerate(ranked, 1):
        avg = total / count
        table.add_row(str(i), desc, f"[red]{_fmt(total)}[/red]", str(count), _fmt(avg))

    console.print()
    console.print(table)
    console.print()


@app.command()
def largest(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    top: int = typer.Option(10, "--top", "-n", help="Number of transactions to show"),
):
    """Show the largest individual transactions."""
    txns = _load(files)
    if not txns:
        raise typer.Exit(1)

    big = largest_transactions(txns, n=top)

    table = Table(title=f"Top {top} Largest Transactions", box=box.ROUNDED)
    table.add_column("Date")
    table.add_column("Description")
    table.add_column("Category")
    table.add_column("Amount", justify="right")

    for t in big:
        table.add_row(t.date.strftime("%d %b %Y"), t.description, t.category or "—", f"[red]{_fmt(abs(t.amount))}[/red]")

    console.print()
    console.print(table)
    console.print()


@app.command()
def transactions(
    files: list[Path] = typer.Argument(..., help="CSV bank statement file(s)"),
    month: Optional[str] = typer.Option(None, "--month", "-m", help="Filter to a specific month (YYYY-MM)"),
    category: Optional[str] = typer.Option(None, "--category", "-c", help="Filter by category name"),
    search: Optional[str] = typer.Option(None, "--search", "-s", help="Search description"),
    debits_only: bool = typer.Option(False, "--debits", help="Show debits only"),
    credits_only: bool = typer.Option(False, "--credits", help="Show credits only"),
):
    """List individual transactions with optional filters."""
    txns = _load(files)
    if not txns:
        raise typer.Exit(1)

    if month:
        txns = [t for t in txns if t.date.strftime("%Y-%m") == month]
    if category:
        txns = [t for t in txns if (t.category or "").lower() == category.lower()]
    if search:
        txns = [t for t in txns if search.lower() in t.description.lower()]
    if debits_only:
        txns = [t for t in txns if t.is_debit]
    if credits_only:
        txns = [t for t in txns if t.is_credit]

    table = Table(title=f"Transactions ({len(txns)})", box=box.ROUNDED)
    table.add_column("Date")
    table.add_column("Description")
    table.add_column("Category")
    table.add_column("Amount", justify="right")
    if any(t.balance is not None for t in txns):
        table.add_column("Balance", justify="right")

    show_balance = any(t.balance is not None for t in txns)
    for t in txns:
        colour = "red" if t.is_debit else "green"
        amt_str = f"[{colour}]{_fmt(t.amount)}[/{colour}]"
        row = [t.date.strftime("%d %b %Y"), t.description, t.category or "—", amt_str]
        if show_balance:
            row.append(_fmt(t.balance) if t.balance is not None else "—")
        table.add_row(*row)

    console.print()
    console.print(table)
    console.print()


@app.command(name="add-rule")
def add_rule(
    category: str = typer.Argument(..., help="Category name (e.g. 'Groceries')"),
    keywords: list[str] = typer.Argument(..., help="Keywords to match in transaction description"),
):
    """Add a custom categorisation rule.

    Example: budget add-rule 'Gym' puregym gymshark
    """
    save_custom_rule(category, keywords)
    console.print(f"[green]✓[/green] Saved rule: [bold]{category}[/bold] ← {', '.join(keywords)}")


def main():
    app()


if __name__ == "__main__":
    main()
