# docx_annotation
Extract highlights from docx files and store as JSON (useful for ML annotations)


## How to use this

This is written in the partial expectation that some people may have decided to back their text extraction with an SQLite DB. 
If this is not your case, you can use it as follows:

python interpret-annotated-files.py -d ../my_annotated_docx_directory/  -o ./json -c schema.txt

and it will scan through all of your docx files looking for colour highlights, use the table you've given to figure out which colours imply which highlights, and store the result. 

## Schemas

To use this usefully, you need a schema (e.g. what colour corresponds to which annotation?)

An example of such a schema is:

BRIGHT_GREEN,APP_DETERMINED_ALLOWED
PINK,FT_OVERTURNED
BLUE,APP_DETERMINED_DISMISSED
TURQUOISE,APP_RESUMED
YELLOW,FT_UPHELD
RED,APP_WITHDRAWN
GRAY_25,APP_REMITTED

Place your version of this into a file, such as for example schema.txt, and reference it on the command line with -c. 

# How to use this on database-backed projects

If this is your case, you can specify the database with the -b flag. 
Further documentation on this case may be added later.
