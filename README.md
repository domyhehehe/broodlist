# Bloodline Pedigree Generator

Create Excel and image pedigrees from `bloodline.csv` by entering a PrimaryKey.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

```powershell
python src\make_pedigree.py
```

For the circular image chart:

```powershell
python src\make_pedigree_image.py
```

To generate a rotating GIF from the circle image:

```powershell
python src\make_circle_gif.py <PK>_circle.png
```

You can also pass a PK and output path:

```powershell
python src\make_pedigree.py PC_INB --out out\PC_INB.xlsx
```

## Output

The Excel file contains a right-expanding pedigree chart:

- Self at the left, ancestors expanding to the right
- Sire on the upper branch, dam on the lower branch
- Each ancestor displayed as `Horse Name Year (PrimaryKey)`
- You will be prompted for max generations (default: 5)
- Inbreeding summary is written to the top row; inbred ancestors get a red border

The image file contains a circular pedigree chart:

- Sire side on the left half, dam side on the right half
- Inner rings are closer generations; center shows the target horse
- Inbreeding summary is written above the chart; inbred ancestors get a red outline
