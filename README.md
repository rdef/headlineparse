# headlineparse
A Python script that iterates over RTF files exported in the a Factivalike format, compiles them, and spits them out in pandas a DataFrame with a sentiment analysis scoring of the headlines.

# Install
Install manually.

1. Download headlineparse.py and requirements.txt
2. From the current directory, run `pip install -r requirements.txt` in your terminal

# Running headlineparse
`headlineparse` can be run directly from the terminal, from a directory that contains `headlineparse.py` as well as one or more appropriately formatted RTF files in the directory or subdirectories:

  `python -m headlineparse`

`headlineparse` runs using a series of class objects that can be imported, allowing greater introspection and comparison of objects.

# Discussion and example
A short explanation of the tool's origins is available [here](https://robbiefordyce.com/2024/04/26/factivaparse-children-and-media-use/).

# Citation
If you use the tool in your research, please cite the copy hosted by Zenodo, as it contains the tool's DOI.

&emsp;Fordyce, R. (2024). headlineparse (0.1). _Zenodo._ https://doi.org/10.5281/zenodo.11069624
