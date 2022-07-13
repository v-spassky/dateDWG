## Purpose

When dealing with large CAD project it is crucial to make sure that all the drawings are up-to-date before launching into production. 
The script is intended to check for it.

## Usage

Clone git repo:

`git clone https://github.com/v-spassky/dateDWG.git`

Initialize virtual environment and eneter it:

`virtualenv -p python3 .`
`source bin/activate`

Install dependencies:

`pip install -r requirements.txt`

From the project folder run the following command:

```
.../dateDWG$ python datedwg.py \
    --workbook <absolute path to the .xlsx spreadsheet> \
    --sheet <sheet name> \
    --target-column <column in which parts` and assemblies` names reside> \
    --result-column <column to write conclusion to> \
    --directory <directory where to search for assemblies/parts/drawinds files>
```

## Example

```
.../dateDWG$ python datedwg.py \
    --workbook /mnt/d/registries/projectAA53.xlsx \
    --sheet Sheet1 \
    --target-column A \
    --result-column B \
    --directory /mnt/d/projectAA53
```
