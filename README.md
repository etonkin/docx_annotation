# docx_annotation
Extract highlights from docx files and store as JSON (useful for ML annotations)

## The point of this software

Sometimes you want to visually mark-up bits of a text file for some reason or another. There are some very complete interfaces permitting this sort of annotation (GATE, for example, has a solution for this). However, for some purposes, it is useful to have a solution to this that involves less overhead in terms of installation and training. One 'discount' solution is making use of existing applications. In this case, we use MS Word to add highlights to text on the basis of an annotation schema we've selected for ourselves ('PINK equals SARCASM, TURQUOISE equals IRONY'). We then need to extract these highlights.

This software searches through for highlights, tries to apply a known schema and dumps the highlights it finds to JSON. 

## How to use this

This is written in the partial expectation that some people may have decided to back their text extraction with an SQLite DB (a more complicated configuration). If this is not your case, you can use it as follows:

```python interpret-annotated-files.py -d ../my_annotated_docx_directory/  -o ./json -c schema.txt```

and it will scan through all of the docx files in `my_annotated_docx_directory` looking for colour highlights, apply the table you've given to figure out which colours imply which highlights, and store the result. 

## Schemas

To use this usefully, you need a schema (e.g. what colour corresponds to which annotation?)

An example of such a schema is:

```
BRIGHT_GREEN,APP_DETERMINED_ALLOWED
PINK,FT_OVERTURNED
BLUE,APP_DETERMINED_DISMISSED
TURQUOISE,APP_RESUMED
YELLOW,FT_UPHELD
RED,APP_WITHDRAWN
GRAY_25,APP_REMITTED
```

Place your version of this into a file, such as for example schema.txt, and reference it on the command line with `-c`. 

# How to use this on database-backed projects

If this is your case, you can specify the database with the -b flag. 
Further documentation on this case may be added later.
