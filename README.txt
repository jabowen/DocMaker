Summary: This generates a document by combining sections eyou've written and replacing 
varaibles to personalize it for the opportunity at hand. 

#Inputs:
    1. A series of files containing the paragraphs you want combined
    written in the proper format
    2. The filepath leading to these files, which is added to the main
    in the varaible "path"
    3. (Optional) A job file

Output: 
    A microsoft word document named "{Doc Name}". 
    OR
    A text file named "{Doc Name}".

    You will be prompted for {Doc Name} in the command line.

Packages:
    Must install "python-docx"

Rules: 
    1. All data files must be in the same folder. The output will appear in a folder called 
    docs one layer back up the path from this folder
    2. Selectable sections must start with '#' and then a number, followed by '-' and then the
    text. If a file only has one section, it can contain only text
    3. Varaibles start with '{', followed by the name, and end with '}'. The name should be
    descriptive, becuase it will be used to prompt you on what it should be replaced with . Varaibles
    with the same name will only be asked for once
    4.Newlines will be ignored, so use '%' if you want a newline mid section. Newlines will be
    added automatically to the end of each section
    5.Comments go inside '[' and ']' respectivly, and will not appear in the documents
    6.Special characters (#, -, [, ], {, }, \, %) can be used normally if they are preceded by '\'
    7. Tab is '/t' 

Job File Rules:
    This should be stored in the data folder
    First line must be in the form of [file.txt,file2.txt,...]
    Other lines must be in format of VarName:VarValue, with each pair seperated by Newlines
    Speical varNames are {"Word?":y if you want microsoft word doc, n if not
                      "Doc Name": the name of the document
                      "Title": the title
                      name of file: selection from file}
    These will be called in every job file

    
