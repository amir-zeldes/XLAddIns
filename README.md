# XLAddIns
Repository for Excel add-ins used for corpus annotation.

Currently available:
==============
ValidateAnnotation.xla - uses a schema file to check for each column header (e.g. pos, word, sentence...):
* Cells in columns with headers "word", "pos" have exact same spans (word=pos)
* "pos" may only contain set values (values:pos(NOUN,VERB,null))
* A worksheet "meta" must contain annotation names in first column (require:meta(title,author,...))
* Span of "sentence" must include all rows of overlapping "word" cell (sentence>word)
* You may also use either/or in headers like this: sent|sent_n>translation

MiscAnno.xla - miscellaneous macros to facilitate annotation tasks
* Auto merge down - merge cells in selection with any empty cells under them
* Stretch into gap - stretches all cell to merge back into preceding cells if they are empty 
   (useful if you insert a line to add a token and want to stretch all span annotations to cover the new tokens)
* Merge with content - merges selected cells and concatenates their contents into the new cell
* Delete line - programmatically delete the line that the cursor is own without marking it
* Insert line - programmatically insert a new line at the cursor position
* Fill zeros where cells in selection are empty (useful for sparse matrices in cross tabulation)

exmaralda_io_0.9.8.1.xla - Legacy version of corpus linguistics format converter for Windows Excel
* Should hopefully be replaced soon by updated multiformat converter version
* See documentation in README_exmaralda_io_0.9.8.1.pdf

UnderOverUse_V1.2.xla - Legacy version of a conditional formatting script for number comparisons
* Highlights values of cells in warm/cold colors by deviation from control column values
* Thresholds are configurable (see underuse_overuse.pdf)
* Mostly reproduceable using conditional formatting in newer versions of Excel

Schemas for ValidateAnnotation.xla (Coptic SCRIPTORIUM only)
* validate_ms.txt validates files created from Coptic manuscripts according to Coptic SCRIPTORIUM standards
* validate_bible.txt validates files created from biblical editions (not manuscripts) such as Sahidica according to Coptic SCRIPTORIUM standards
