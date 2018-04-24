# transcript2haml
Convert transcript received from word document to HAML with spreadsheet + macros.

## To Use

* Create a new Excel spreadsheet.
* Open the VBA console (Alt+F11)
* Create a new module and copy in the contents of main.vba
* Create a sheet where you can paste the contents of a transcript.  Add named range 'contents' in cell A1 of this sheet.
* Setup an area for settings on the main worksheet.  Add a named range for 'row_class', 'scolumn', 'sheader', 'tcolumn', 'paragraph', and 'out_file'.  Optionally add a buttons to run the 'do_all' macro.
* Copy paste your transcript to the contents worksheet.
* Update configuration by entering the appropriate HAML codes you'd like to appear in the output and set an output file.
* Run do_all

For more information on what this is for, read the relevant [blog post](http://mikekling.com/translate-interview-transcript-to-haml-with-excel/).
