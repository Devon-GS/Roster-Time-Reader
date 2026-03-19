# Roster-Time-Reader

Calculates fortnight working hours of employees based on company's specific roster. 
Puts roster times and hours worked in to excel workbook.

### <ins>Fixed Issues:</ins>

- 08 March 2024     [Patch001]  - Fixed issues with hour calculation
- 05 November 2025  [Patch002]  - Add error handling 
- 11 November 2025  [Patch003]  - Add formating of the excel sheet [Rewrite finished] 
- 19 November 2025  [Patch004]  - Changed paths for Excel, Database and Icons for production 
- 09 March 2026  	[Patch005]  - Fixed while app was running was unable to save excel workbook
- 19 March 2026  	[Patch006]  - Fixed blank cell throws error
- 19 March 2026  	[Patch007]  - Add total hours. Update formating of excel workbook

<!-- ### <ins>Currently Working On:</ins>

- N/A -->

### <ins>Installing and Running Program:</ins>

Install environment

```python
python -m venv <name>
```

Install poetry

```python
pip install poetry
``` 

Install dependencies

```python
poetry install
```

Run Script

```python
python main.py