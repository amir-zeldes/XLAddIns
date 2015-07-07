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

