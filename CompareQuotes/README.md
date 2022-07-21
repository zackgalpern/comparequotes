# Python Quote Comparison

## IMPORTANT
### Rename new quote and old quote to NEW.xlsx and OLD.xlsx, respectively

## Summary
This Python script automates the comparison of two Quotes, specifically
Manufacturer Part # and Qty. It will place an outer join of the results into 
a text file for easy comparison.

## How to use
### Install Packages
First, test if you have Pandas installed by opening up Terminal (Mac) or Command Prompt (Windows)
and typing:
#### Mac
```pip3 show pandas```

#### Windows
```py -m pip show pandas```

If Pandas is installed, move on. If not, install it:

#### Mac
```pip3 install pandas```

#### Windows
```py -m pip install pandas```

### Preparing the documents
Both quote Excel documents will contain multiple sheets (i.e. MDF, IDF, etc).
To use this script, select the sheet you want to compare. Right click the sheet
on the sheet menu on the bottom of the Excel documents and click on 'Move or Copy'.
In the "To book:" drop down menu, select "(new book)". Also select "Create a copy".
The newly created Excel document will be the correct one to use. Rename it to either
OLD.xlsx or NEW.xlsx depending on if it is the old or new quote.

### Running the script
Click into the CompareQuotes project folder until you reach the Quotes folder.
The Quotes folder is where you will place both NEW.xlsx and OLD.xlsx. Open up Terminal (Mac)
or Command Prompt (Windows) and cd into the project folder:

```cd CompareQuotes/CompareQuotes```

Then run the command:

#### Mac
```python3 CompareQuotes.py```

#### Windows
```py CompareQuotes.py```

### Obtaining the results
Inside the CompareQuotes project folder, there will be a folder called Results. 
After running the script, a text file will be created in here called results.txt,
and it will include the sheet name along with the comparison of the old and new quote
manufacturer part #s and quantities.

## Troubleshooting
Here are some common errors related to 'openpyxl'
- > ImportError: Missing optional dependency 'openpyxl'.  Use pip or conda to install openpyxl.
- > ModuleNotFoundError: No module named 'openpyxl'

To fix, run
#### Mac
```python3 -m pip install openpyxl```

#### Windows
```py -m pip install openpyxl```

If you have any questions, contact me on Teams, at zachary.galpern@universalorlando.com, or my phone number at (954) 646-2929