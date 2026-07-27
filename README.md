# AI Journal Entry Generator for Excel

Convert natural-language business transactions into properly formatted double-entry journal entries — automatically.

Type a transaction the way you'd describe it to a person ("Purchased furniture for ₹50,000 by cheque"), and this tool reads it straight from an Excel sheet, uses an LLM to classify the debit account, credit account, and amount according to real accounting rules (Personal / Real / Nominal account classification, golden rules of debit and credit), and writes back a fully formatted, styled journal in a new Excel workbook.

Built as a learning project to combine accounting fundamentals with AI and Python — not just an AI wrapper, but a tool designed around the actual golden rules of double-entry bookkeeping.

## Example

**Input (`data.xlsx`):**

| Date        | Transaction                               |
| ----------- | ------------------------------------------ |
| 27-Jul-2026 | Purchased furniture for ₹50,000 by cheque  |
| 27-Jul-2026 | Paid salary ₹30,000 in cash                |
| 28-Jul-2026 | Received rent ₹15,000 by bank              |

**Output (`output.xlsx`):**

| Date        | Particulars              | Debit Amount (Rs.) | Credit Amount (Rs.) |
| ----------- | ------------------------- | ------------------: | --------------------: |
| 27-Jul-2026 | Furniture A/C Dr.        |              50,000 |                        |
|             | To Bank A/C                |                      |                50,000 |
|             | *(Being furniture purchased by cheque)* |     |                        |
| 27-Jul-2026 | Salary A/C Dr.            |              30,000 |                        |
|             | To Cash A/C                 |                      |                30,000 |
|             | *(Being salary paid in cash)* |          |                        |
|             | **Total**                 |         **XX,XXX**  |          **XX,XXX**   |

The header row is highlighted, each entry's narration line is underlined to visually separate transactions, and the total row is bordered — with debit and credit totals matching, as they should for any balanced set of journal entries.

## Features

- Reads transactions directly from an Excel input file
- Uses an LLM (via the Groq API) to classify each transaction into a debit account, credit account, and amount
- Applies real double-entry accounting logic — grounded in the Personal/Real/Nominal account classification and the golden rules of debit and credit, not just a generic AI guess
- Auto-generates narration ("Being ___") for each entry
- Outputs a clean, styled Excel workbook: colored header, underlined narration rows, bordered totals
- Debit/credit totals are computed automatically — a built-in sanity check, since a correct set of journal entries always balances

## How it works

1. **Read** — `get_transactions()` opens the input workbook with `openpyxl` and returns each row as a `{"date": ..., "transaction": ...}` dict.
2. **Classify** — `analyze_entries()` sends each transaction to an LLM (Groq, `llama-3.3-70b-versatile`) with a system prompt encoding real accounting rules, and gets back structured JSON: `{"debit_account": ..., "credit_account": ..., "amount": ..., "narration": ...}`.
3. **Format** — the two results are paired up (`zip()`) and turned into the standard three-line double-entry format: the debit line, the "To ___" credit line, and the narration line.
4. **Write** — `write_journal()` writes everything to a new `output.xlsx`, applying header styling, narration underlines, and total borders along the way.

## Tech stack

- **Python**
- **openpyxl** — reading and writing Excel files, including styling (fonts, fills, borders)
- **Groq API** (`llama-3.3-70b-versatile`) — natural language transaction understanding, returned as structured JSON
- **python-dotenv** — loading the Groq API key from a local `.env` file (never committed)

## Setup

```bash
pip install openpyxl groq python-dotenv
```

Create a `.env` file in the project root:

```
GROQ_API_KEY=your_key_here
```

Place your input transactions in `data.xlsx` (columns: `Date`, `Transaction`), then run:

```bash
python journal_entry.py
```

Output is written to `output.xlsx`.

## Current scope (v1)

This first version is intentionally limited to **simple, two-account transactions** — one debit account, one credit account, one amount. Advanced accounting scenarios are deliberately excluded until the underlying concepts are well understood and the core pipeline is solid.

## Roadmap

- [ ] Compound entries — multiple debit and/or credit accounts in a single transaction
- [ ] Malformed-AI-output handling (graceful fallback if the model returns invalid JSON)
- [ ] Ledger posting
- [ ] Trial Balance generation
- [ ] Profit & Loss Account
- [ ] Balance Sheet generation
- [ ] GST handling
- [ ] Invoice and receipt processing
- [ ] PDF report generation
- [ ] Desktop GUI
- [ ] Web application with user authentication

## Why this project

This combines accounting knowledge, AI, and Python to automate a real bookkeeping task — while still respecting the fundamentals of double-entry bookkeeping rather than treating the AI as a black box. It's designed to grow: as the underlying accounting knowledge deepens, so does the tool, from a journal entry generator toward a complete accounting automation system.
