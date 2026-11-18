# Supplier Product Finder

Finds product URLs on supplier/e-commerce sites for spreadsheet rows of **brand + model/part number** — and only returns a URL on an *exact verified* match.

## Method

Written for sites like HD Supply, Zoro and Ferguson (hundreds of rows per batch):

1. Search the supplier's own search endpoint (or a search engine with a `site:` filter when the site blocks direct access) with ordered query passes: `brand + model`, then `brand + supplier item #`, then the item number alone.
2. Filter candidates by **exact** manufacturer part-number equality — a name similarity is never enough; bare-number models are handled carefully because they false-match other brands' part numbers.
3. Verify the surviving candidate on its product page (schema.org JSON-LD: brand + MFG# must match the row) before writing the URL.
4. Brand-alias aware — e.g. a sheet brand of ELECTROLUX also accepts the exact model listed under its owned brand Frigidaire.

Rows that can't be verified are marked Not Found or left Unchecked for a retry pass — never padded with a plausible-looking link.

## Stack

Python · requests · openpyxl · JSON-LD parsing
