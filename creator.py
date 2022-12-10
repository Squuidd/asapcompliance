import safety_program_creator as spc
import os 
from os.path import isfile, join

chosen_files_path = "Chosen Programs"
# onlyfiles = [f for f in listdir(chosen_files_path) if isfile(join(chosen_files_path, f))]

chosen_files = [os.path.join(r,file) for r,d,f in os.walk(chosen_files_path) for file in f]
company_name = input("Company Name: ")
choice = input("1 Safety Manual\n2 Safety Programs and PDFs\n3 Both\n")

def generate_file(bytes, extension):
    with open(f"Output\\{company_name}_manual.{extension}", "wb") as f:
        f.write(bytes)

if choice == '1':
    manual_bytes = spc.create_manual(file=spc.findPath("safety_manual.docx"), safety_documents=chosen_files, company_name=company_name)
    generate_file(manual_bytes, "docx")
elif choice == '2':
    program_bytes = spc.create_program(files=chosen_files, company_name=company_name)
    generate_file(program_bytes, "zip")
else: 
    program_bytes = spc.create_program(files=chosen_files, company_name=company_name)
    generate_file(program_bytes, "zip")
    manual_bytes = spc.create_manual(file=spc.findPath("safety_manual.docx"), safety_documents=chosen_files, company_name=company_name)
    generate_file(manual_bytes, "docx")








