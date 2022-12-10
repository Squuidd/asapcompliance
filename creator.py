import safety_program_creator as spc
from os import listdir
from os.path import isfile, join

chosen_files_path = "Chosen Programs"
onlyfiles = [f for f in listdir(chosen_files_path) if isfile(join(chosen_files_path, f))]

manual_bytes = spc.create_manual(file=spc.findPath("safety_manual.docx"), safety_documents=[f"Chosen Programs\\{onlyfiles[0]}"], company_name="Test")

with open("Chosen Programs\\new.docx", "wb") as f:
    f.write(manual_bytes)

#program_bytes = spc.create_program([f"Chosen Programs\\{onlyfiles[0]}"], company_name="Test")
# print(manual_bytes)

# with open(manual_bytes, mode="rb") as f:
#     file = f.read()

# with spc.Tempdoc(file) as doc:
#     bytes = doc.save(".docx")

# with open("Chosen Programs\\new.docx", "wb") as binary_file:
#     binary_file.write(file)


