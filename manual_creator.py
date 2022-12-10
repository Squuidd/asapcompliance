import safety_program_creator as spc
import os 

while True:
    chosen_files_path = "Chosen Programs"
    chosen_files = [os.path.join(r,file) for r,d,f in os.walk(chosen_files_path) for file in f]
    company_name = input("Company Name: ")
    choice = input("1 Safety Manual\n2 Safety Programs and PDFs\n3 Both\n")

    def generate_file(bytes, extension, type):
        with open(f"Output\\{company_name}_{type}.{extension}", "wb") as f:
            f.write(bytes)



    filename = "safety_manual.docx"
    # if '_MEIPASS2' in os.environ:
    #     filename = os.path.join(os.environ['_MEIPASS2'], filename)



    if choice == '1':
        try:
            manual_bytes = spc.create_manual(filename, safety_documents=chosen_files, company_name=company_name)
            generate_file(manual_bytes, "docx", "manual")
        except Exception as e:
            print(e)

    elif choice == '2':
        program_bytes = spc.create_program(files=chosen_files, company_name=company_name)
        generate_file(program_bytes, "zip", "programs")

    else: 
        program_bytes = spc.create_program(files=chosen_files, company_name=company_name)
        generate_file(program_bytes, "zip", "programs")
        manual_bytes = spc.create_manual("safety_manual.docx", safety_documents=chosen_files, company_name=company_name)
        generate_file(manual_bytes, "docx", "manual")

    print("Check Output folder.\n\n")
    







