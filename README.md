# docx-merge
A Python application to perform merges on DOCX format files.

Uses DOCX file templates (from Word, or exported from Google Docs) with mustaces-style variables (i.e. {{var}}).

Can handle Excel / CSV data sources. Can also handle iCal files for calendar data.

Can do standard merges (one document clone with variables replaced per line in a spreadsheet), calendar merges (one document clone per month or week in an iCal file) and sticker merges (one value per item on a page).
