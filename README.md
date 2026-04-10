# Jeopardy PowerPoint Generator

Generate a fully-navigable Jeopardy-style PowerPoint (`.pptx`) deck with a single Python command.

## Features

- **Main board slide** — 6 × 5 Jeopardy grid with Jeopardy-blue background, white category headers, and gold dollar-value cells
- **30 Clue slides** — one per cell, with a header and a blank body ready for you to fill in
- **30 Answer slides** — linked from each clue slide, also blank for editing
- **Full hyperlink navigation** inside the deck:

  ```
  Board cell  →  Clue slide  →  Answer slide  →  Board
  ```

- **Configurable round type**: Single ($200–$1000), Double ($400–$2000), Triple ($600–$3000)
- **Custom categories** and **custom dollar values** via CLI flags

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

### Basic (all defaults)

```bash
python jeopardy.py
```

Produces `jeopardy.pptx` with six placeholder categories and Single Jeopardy values.

### Custom categories

```bash
python jeopardy.py --categories "Science" "History" "Math" "Pop Music" "Sports" "Geography"
```

> Wrap multi-word category names in quotes.

### Round type

```bash
python jeopardy.py --round double
python jeopardy.py --round triple
```

| Round  | Dollar values                        |
|--------|--------------------------------------|
| single | $200, $400, $600, $800, $1,000       |
| double | $400, $800, $1,200, $1,600, $2,000   |
| triple | $600, $1,200, $1,800, $2,400, $3,000 |

### Custom dollar values

```bash
python jeopardy.py --values 100 200 300 400 500
```

> `--values` overrides `--round`.

### Custom output filename

```bash
python jeopardy.py --output my_game.pptx
```

### Full example

```bash
python jeopardy.py \
  --categories "Science" "History" "Math" "Pop Music" "Sports" "Geography" \
  --round double \
  --output double_jeopardy.pptx
```

---

## Editing the Deck

After generating the `.pptx`:

1. Open it in **Microsoft PowerPoint** or **Google Slides**.
2. Navigate to each **Clue slide** (slides 2, 4, 6, …) and replace `[ Enter Clue Here ]` with your clue text.
3. Navigate to each **Answer slide** (slides 3, 5, 7, …) and replace `[ Enter Answer Here ]` with the correct answer.
4. Save and present — the hyperlink buttons work in presentation mode.

---

## Slide Structure

| Slide(s) | Content |
|----------|---------|
| 1 | Main Jeopardy board |
| 2, 3 | Clue & Answer — Category 1 / Value 1 |
| 4, 5 | Clue & Answer — Category 1 / Value 2 |
| … | … |
| 60, 61 | Clue & Answer — Category 6 / Value 5 |

---

## Requirements

- Python 3.7+
- [`python-pptx`](https://python-pptx.readthedocs.io/) ≥ 0.6.21
