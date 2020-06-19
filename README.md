### Overview
Python Script to automate building and importing word documents in PDF format. Streamlined process by reading in simple text files or user inputs to build the documents   


### Required Libary
```sh
$ pip install docxtpl
$ pip install comtypes
```

### Prior Set Up
- Replace sections or words with {{ var_name }} in your template word file
- In context.txt put all the var_names in a dictionary format. ex
    ```sh
    {'date': '', 'first_name': '', 'last_name': ''}
    ```
- In generateDoc.py, update line 27 to the name of your template file ex.
    ```sh
    doc = DocxTemplate('new-Coverletter.docx')
    ```
- Make sure all of your template files, context.txt, and generateDoc.py are in the same directory


### Run Instruction
- Run generateDoc.py
    ```sh
    $ python generateDoc.py
    ```
- Add texts asked by the script ex.
    ```sh
    date: May 16th, 2020
    first_name: John
    last_name: Doe
    ```
- To grab inputs from text file put name of the file ex.
    ```sh
    paragraph1: para.txt
    ```
- At last step, the script asks for new name for the document ex.
    ```sh
    new file name: new_name.py
    ```