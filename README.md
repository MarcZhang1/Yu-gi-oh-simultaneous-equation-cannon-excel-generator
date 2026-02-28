# Yu-Gi-Oh! Simultaneous Equation Cannon Excel Generator

Python script that generates an Excel spreadsheet listing every valid play for the
[Simultaneous Equation Cannon](https://www.db.yugioh-card.com/yugiohdb/card_search.action?ope=2&cid=19921) trap card.

## Install

```bash
pip install openpyxl
```

## Usage

```
python generate.py <xyz_min> <xyz_max> <fusion_min> <fusion_max>
```

```bash
python generate.py 2 6 1 5   # -> results/sec xyz2-6 fusion1-5.xlsx
python generate.py 3 6 1 6   # -> results/sec xyz3-6 fusion1-6.xlsx
```

The output is saved to the `results` folder.
