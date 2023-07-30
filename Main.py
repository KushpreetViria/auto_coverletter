from PrintColors import ColorPrint as p
from CoverEdit import DocxEditor as editor
import argparse
import traceback
import sys
import os


def setupArgs():
    parser = argparse.ArgumentParser(
        description="Script for automating coverletter creation by editing just company specific things in a template document")
    parser.add_argument("--template", help="Coverletter template to populate", required=False)
    parser.add_argument("--streetaddress", help="The street address of company", required=False)
    parser.add_argument("--city", help="The city, province, country of company", required=False)
    parser.add_argument("--company_full", help="The full company name", required=False)
    parser.add_argument("--company_short", help="The short company name", required=False)
    parser.add_argument("--position", help="The job position", required=False)
    return parser.parse_args()

def createEditorData(args):
    def getInput(string):
        user_input = input(string)
        user_input = None if not user_input.strip() else user_input
        return user_input

    if (args.streetaddress == None):    args.streetaddress = getInput("Street address: ")
    if (args.city == None):             args.city = getInput("City address: ")
    if (args.company_full == None):     args.company_full = getInput("Company full name: ")

    while (args.company_short == None): args.company_short = getInput("Company short nnme: ")
    while (args.position == None):      args.position = getInput("Position: ")
    return args


if __name__ == "__main__":
    file = "COVERLETTER_GENERAL.docx" # default template

    args = setupArgs()
    if args.template and not os.path.exists(args.template):
        p.print_error(f"Invalid arguments provided. No such file ({args.template})")
        sys.exit(0)
    elif args.template:
        file = args.template
        del args.template
    args = createEditorData(args)

    try:
        e = editor(file)
        p.print_color("Editing document...",styles=["bold","italic"])
        e.populate_data(street=args.streetaddress,city=args.city,company_full=args.company_full,
            company_short=args.company_short,position=args.position)

        p.print_color("Saving document...",styles=["bold","italic"])
        doc = e.save(output_path=f"output/{args.company_short}_coverletter")
        p.print_color("Converting to PDF...",styles=["bold","italic"])
        pdf_doc = editor.save_as_pdf(doc)

        if editor.get_number_of_pdf_pages(pdf_doc) > 1:
            os.remove(pdf_doc)
            p.print_color(
                f"New doc ({doc}) is multiple pages long. Edit manually.",
                color="yellow",
                styles=["bold", "underline"],
            )
            p.print_color("Exiting.")
            exit(0)
        else:
            os.remove(doc)

        full_path = os.path.join(os.getcwd(), pdf_doc).replace('\\', '/')
        p.print_color(f"Finished Succesfully. Output = {full_path}", color="green")

    except Exception as e:
        traceback.print_exc()
        p.print_error(f" error occurred: {e}")