import os
import docx2txt
import re

#setting our source and text file directories
source_directory = os.path.join(os.getcwd(), "source")
text_directory = os.path.join(os.getcwd(), "text_files")

for process_file in os.listdir(source_directory):
    #solving OSX issues with autogeneration of .DS_Store
    if process_file == ".DS_Store":
        continue

    #saving filename and extension separately for use in loop
    file, extension = os.path.splitext(process_file)

    # We create a new text file name by concatenating the .txt extension to file
    dest_file_path = file + '.txt'

    #extract text from the file
    try:
        content = docx2txt.process(os.path.join(source_directory, process_file))
    except Exception:
        input(process_file)
    # We create and open the new and we prepare to write the Binary Data which is represented by the wb - Write Binary
    write_text_file = open(os.path.join(text_directory, dest_file_path), "wb")

    #write the content and close the newly created file
    write_text_file.write(content)
    write_text_file.close()
# ###pasa a csv
with open('salida.csv','wb') as file:
    file.write(bytes("Nombre; Apellido; Region; Telefono; Email; Institucion; Cargo\n".encode("utf8")))
    for textfile in os.listdir(os.path.join("text_files")):
        if textfile == ".DS_Store":
            continue
        with open(os.path.join("text_files",textfile), "r", encoding = "utf8") as archo:
            texto = []
            try:
                for linea in archo:
                    mat = 0
                    match = re.search(r'Nombre: (\S+)', linea)
                    if match:
                        nombre = match.group(1)
                    match1 = re.search(r'Apellido: (\S+)', linea)
                    if match1:
                        apellido = match1.group(1)
                    match2 = re.search(r'Región:[ ]+([A-zÀ-ú ]+)', linea)
                    if match2:
                        region = match2.group(1)
                    match3 = re.search(r'Tel Móvil \(Código de País - Código de Área - Número\):([0-9\s\-\+]+)', linea)
                    if match3:
                        tel = match3.group(1).strip()
                    mat += 1
                    match4 = re.search(r'Email: (\S+)', linea)
                    if match4:
                        mail = match4.group(1)
                    match5 = re.search(r'Título Académico: [ ]+([A-zÀ-ú ]+)', linea)
                    # if match5:
                    #     titulo = match5.group(1)
                    # else:
                    #     titulo = ""
                    match6 = re.search(r'Lugar de trabajo\/ Institución:[ ]+([A-zÀ-ú ]+)', linea)
                    if match6:
                        institucion = match6.group(1)
                    match7 = re.search(r'Cargo actual: (\S+)', linea)
                    if match7:
                        cargo = match7.group(1)
            except Exception:
                input(textfile)
            if mat == 0:
                tel = ""
            row = ";".join([nombre,apellido,region,tel,mail,institucion,cargo, "\n"])
            # rowb = bytes(row, encoding = "utf-8")
            file.write(row.encode('utf8'))
            # file.write('\n')
