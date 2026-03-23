EU Customs Master Template — README
==================================

Files:
- EU_Customs_Master_Template.xlsx — main workbook (country selector, localization, VAT types)
- ExportToPDF.bas — VBA module to enable a "Print to PDF" macro (optional)

How to enable the PDF macro:
1) Open the workbook and save it as **.xlsm** (macro-enabled).
2) Press ALT+F11 (open VBA editor) → File → Import File… → select **ExportToPDF.bas**.
3) (Optional) Add a shape/button on the "Declaration" sheet and assign macro: **Export_Declaration_and_Items_PDF**.
4) Run the macro to export "Declaration" + "Items" into one PDF.

Localization:
- Labels on 'Declaration' are localized automatically to EN/EE/DE/FR based on the selected Country.
- Change mapping on hidden 'Lists' → CountryLanguageMap. Add/edit translations on hidden 'Translations'.

VAT:
- Default VAT: Standard/Reduced1/Reduced2 by Country (hidden 'Lists' → VAT_Table).
- Each item has "VAT Type" (Standard/Reduced1/Reduced2/Manual). If Manual, fill "Manual VAT (%)".
