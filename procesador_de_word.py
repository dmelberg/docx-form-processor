import os
import docx2txt
import re

######## PART I: GENERATING TEXT FILES ########

#setting our source and text file directories
source_directory = os.path.join(os.getcwd(), "source")
text_directory = os.path.join(os.getcwd(), "text_files")

for process_file in os.listdir(source_directory):
    #solving OSX issues with autogeneration of .DS_Store
    if process_file == ".DS_Store":
        continue

    #saving filename and extension separately for use in loop
    filename, extension = os.path.splitext(process_file)

    # We create a new text file name by concatenating the .txt extension to file
    dest_file_path = filename + '.txt'

    #extract text from the file
    try:
        content = docx2txt.process(os.path.join(source_directory, process_file))
    except Exception:
        #catch file causing error and show me its name before moving on
        input(process_file)
    
    # We create and open the new and we prepare to write the Binary Data which is represented by the wb - Write Binary
    write_text_file = open(os.path.join(text_directory, dest_file_path), "wb")

    #write the content and close the newly created file
    write_text_file.write(content)
    write_text_file.close()

#future refactor: do both parts in single loop writing line by line in csv

########## PART II: TEXT TO CSV ##########
regex = {
    "nombre": r'Nombre: (\S+)',
    "apellido": r'Apellido: (\S+)',
    "region": r'Región:[ ]+([A-zÀ-ú ]+)',
    "tel": r'Tel Móvil \(Código de País - Código de Área - Número\):([0-9\s\-\+]+)',
    "email": r'Email: (\S+)',
    "titulo": r'Título Académico: [ ]+([A-zÀ-ú ]+)',
    "institucion": r'Lugar de trabajo\/ Institución:[ ]+([A-zÀ-ú ]+)',
    "cargo": r'Cargo actual: (\S+)'
}

def get_match(field, line):
    """
    Get regex matches from list above in text files
    """
    match = re.search(field, line)
    if match:
        return match.group(1)

with open('formdata.csv','wb') as csvfile:
    #set csv headers
    csvfile.write(bytes("Nombre; Apellido; Region; Telefono; Email; Institucion; Cargo\n".encode("utf8")))
    #loop over textfiles
    for textfile in os.listdir(os.path.join("text_files")):
        #osx workaround
        if textfile == ".DS_Store":
            continue
        with open(os.path.join("text_files",textfile), "r", encoding = "utf8") as tfile:
            items = []
            #trying to fill as much as possible
            try:
                for line in tfile:
                    for key in regex:
                        items.append(get_match(regex[key], line))
            except Exception:
                input(textfile)
            items.append("\n")
            row = ";".join(items)
            csvfile.write(row.encode('utf8'))
