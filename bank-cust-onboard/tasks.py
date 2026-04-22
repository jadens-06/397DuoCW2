from pathlib import Path
import re

from robocorp import workitems
from robocorp.tasks import get_output_dir, task
from RPA.Excel.Files import Files as Excel


@task
def producer():
    """Split Excel rows into multiple output Work Items for the next step."""
    output = get_output_dir() or Path("output")
    filename = "orders.xlsx"

    for item in workitems.inputs:
        path = item.get_file(filename, output / filename)

        excel = Excel()
        excel.open_workbook(path)
        rows = excel.read_worksheet_as_table(header=True)

        for row in rows:
            payload = {
                "Name": row["Name"],
                "Zip": row["Zip"],
                "Product": row["Item"],
            }
            workitems.outputs.create(payload)


@task
def consumer():
    zip_code_re = r"^\d{5}(-\d{4})?$"

    for item in workitems.inputs:
        try:
            name = item.payload["Name"]
            zipcode = item.payload["Zip"]
            product = item.payload["Product"]

            print(f"Processing order: {name}, {zipcode}, {product}")

            if not re.match(zip_code_re, str(zipcode)):
                raise AssertionError("Invalid ZIP code")

            item.done()

        except AssertionError as err:
            item.fail("BUSINESS", code="INVALID_ORDER", message=str(err))

        except KeyError as err:
            item.fail("APPLICATION", code="MISSING_FIELD", message=str(err))
