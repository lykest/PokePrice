# PokePrice

A utility to take an Excel spreadsheet (database) full of monsters and figure out what their approximate value 
is from scraping a price aggregator with feeds from several auction sites.

The program expects a single-sheet Excel workbook with the first row as labels for the following five fields (columns):

    Name  <Req> |  Type <Opt> |  Marking <Opt> |  EN count <Opt> |  JP count <Opt>

    ex:  Pikachu  |  Electric  |  PROMO  |  2  |  5

The EN and JP count are used as multipliers for the determined average, low, and high values for a given Name.
Only Name is required for practical functionality, as the program assumes EN as 1 if not specified.

Extraneous columns are ignored by the program.  Only the Name (first) column is required.

![PokePrice Results]
(https://github.com/lykest/PokePrice/blob/master/results.png)
