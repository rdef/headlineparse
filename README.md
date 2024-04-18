# factivaparse
A Python script that iterates over RTF files exported from Factiva, compiles them, and spits them out in pandas a DataFrame with a sentiment analysis scoring of the headlines.

# Install
Install manually.

1. Download factivaparse.py and requirements.txt
2. From the current directory, run `pip install -r requirements.txt` in your terminal

# Running factivaparse
factivaparse can be run directly from the terminal, from a folder that contains `factivaparse.py` one or more RTF files exported from Factiva in the directory or subdirectory:
  `python -m factivaparse`

factivaparse runs using a series of class objects that can be imported, allowing greater introspection and comparison of objects.
