import zipfile
from io import BytesIO


def safe_filename(name: str):

    return (
        str(name)
        .replace("/", "-")
        .replace("\\", "-")
        .replace(":", "-")
        .strip()
    )


def build_reports_zip(
    brand_workbooks: dict
):
    """
    brand_workbooks:
        {
            brand_name: {
                "buffer": BytesIO,
                "has_sales": bool
            }
        }
    """

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        for brand, data in brand_workbooks.items():

            buffer = data["buffer"]
            has_sales = data["has_sales"]

            if buffer is None:
                continue

            safe_name = safe_filename(brand)

            if has_sales:
                path = f"Reports/{safe_name}.xlsx"
            else:
                path = f"Reports/Empty Brand Guard/{safe_name}.xlsx"

            zip_file.writestr(path, buffer.getvalue())

    zip_buffer.seek(0)

    return zip_buffer
