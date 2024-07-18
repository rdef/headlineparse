# headlineparse
A Python script that iterates over RTF files exported in the a Factivalike format, compiles them, and spits them out in pandas a DataFrame with a sentiment analysis scoring of the headlines.

# Install
Install manually.

1. Download and unzip the [latest release](https://github.com/rdef/headlineparse/releases/latest) or download `headlineparse.py` and `requirements.txt`
2. Place the downloads in a working directory that contains at least one RTF file of headlines from a news database such as Factiva.
3. From your working directory, run `pip install -r requirements.txt` in your terminal. This ensures that you have the required libraries, but will not install `headlineparse` to your PATH.

A version will eventually be made available on PyPi, and this ReadMe will be updated to reflect that when it happens.

# Running headlineparse
`headlineparse` can be run directly from the terminal within a directory that contains `headlineparse.py` as well as one or more appropriately formatted RTF files in that directory or subdirectories:

&emsp;`python -m headlineparse`

The tool does not currently have a hardcoded limit on the number or size of RTF files that can be scanned, so running this from Home or Documents directory is likely to break or cause issues in some manner and is not recommended.

# Discussion and example
`headlineparse` runs using a series of class objects that can be imported, allowing greater introspection and comparison of objects.

A short explanation of the tool's origins is available [here](https://robbiefordyce.com/2024/04/26/factivaparse-children-and-media-use/).

# Citation
If you use the tool in your research, please cite the copy hosted by Zenodo, as it contains the tool's DOI.

&emsp;Fordyce, R. (2024). headlineparse (0.1). _Zenodo._ https://doi.org/10.5281/zenodo.11069624
