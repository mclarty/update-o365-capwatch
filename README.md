# update-o365-capwatch
Python script to update an Office 365 tenant users with data from Civil Air Patrol CAPWATCH database

## To Use

1. Download/clone the repo
2. Download a CAPWATCH database containing the scope of the users to be managed into a directory called `capwatch` within the repo's root directory
3. Run the Python script (e.g., `python3 update-o365-capwatch.py`)
4. Use the resulting output CSV file as the basis for a PowerShell script to perform a bulk update of user accounts

```
TODO: Write an exemplar PowerShell script
```
