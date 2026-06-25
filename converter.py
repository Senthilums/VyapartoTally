from vyapar_daybook_to_tally_xml import main as run_converter


def generate_xml(
    input_file,
    output_folder,
    generate_masters=True,
    generate_vouchers=True,
    sales=True,
    receipts=True
):
    if not input_file:
        raise ValueError("Please select Vyapar Excel file.")

    if not output_folder:
        raise ValueError("Please select output folder.")

    run_converter(
        input_file=input_file,
        output_dir=output_folder,
    )