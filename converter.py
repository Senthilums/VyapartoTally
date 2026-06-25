from vyapar_daybook_to_tally_xml import MAX_DATA_AGE_DAYS, main as run_converter


def generate_xml(
    input_file,
    output_folder,
    items_input_file=None,
    generate_masters=True,
    generate_vouchers=True,
    generate_combined=True,
    sales=True,
    receipts=True,
    purchases=True,
    payments=True,
    notes=True,
    allow_accounting_only_sales=False,
    max_data_age_days=MAX_DATA_AGE_DAYS,
):
    if not input_file:
        raise ValueError("Please select Vyapar Excel file.")

    if not output_folder:
        raise ValueError("Please select output folder.")

    if not generate_masters and not generate_vouchers:
        raise ValueError("Please select at least one XML file type to generate.")

    included_voucher_types = []
    if sales:
        included_voucher_types.append("sales")
    if receipts:
        included_voucher_types.append("receipt")
    if purchases:
        included_voucher_types.append("purchase")
    if payments:
        included_voucher_types.append("payment")
    if notes:
        included_voucher_types.extend(["credit_note", "debit_note"])

    if generate_vouchers and not included_voucher_types:
        raise ValueError("Please select at least one voucher type.")

    return run_converter(
        input_file=input_file,
        items_input_file=items_input_file or None,
        output_dir=output_folder,
        generate_masters=generate_masters,
        generate_vouchers=generate_vouchers,
        generate_combined=generate_combined,
        included_voucher_types=included_voucher_types,
        allow_accounting_only_sales=allow_accounting_only_sales,
        max_data_age_days=max_data_age_days,
    )
