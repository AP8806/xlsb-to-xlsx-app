import io
import platform
import tempfile
from pathlib import Path
import pandas as pd

XLSX_FORMAT = 51  # xlOpenXMLWorkbook


def describe(exc: Exception) -> str:
    info = getattr(exc, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return f"{info[2]} (code {info[-1]})"
    return str(exc)


def open_workbook(excel, path: Path):
    attempts = [
        dict(Filename=str(path), UpdateLinks=0, ReadOnly=True, Password="unused-password"),
        dict(Filename=str(path), UpdateLinks=0, ReadOnly=True, Password="unused-password", CorruptLoad=1),
    ]
    last_error = None
    for kwargs in attempts:
        try:
            return excel.Workbooks.Open(**kwargs)
        except Exception as exc:
            last_error = exc
    raise last_error


def convert_via_python_engine(file_bytes: bytes) -> bytes:
    """Cloud server var (bina MS Excel) .xlsb che sagle sheets .xlsx madhe convert karne"""
    xlsb_io = io.BytesIO(file_bytes)
    output_io = io.BytesIO()
    
    excel_file = pd.ExcelFile(xlsb_io, engine="pyxlsb")
    with pd.ExcelWriter(output_io, engine="openpyxl") as writer:
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            
    return output_io.getvalue()


def convert_uploaded_files(uploaded_files, progress_callback=None):
    converted_files = []  # (filename.xlsx, bytes)
    failed_files = []     # (filename.xlsb, error_message)
    total = len(uploaded_files)

    # Check kara ki Windows + MS Excel uplabdha ahe ka
    use_com = False
    excel = None
    pythoncom = None

    if platform.system() == "Windows":
        try:
            import pythoncom as _pythoncom
            import win32com.client
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            use_com = True
        except Exception:
            use_com = False

    try:
        with tempfile.TemporaryDirectory(prefix="xlsb_app_") as work:
            work = Path(work)

            for number, up_file in enumerate(uploaded_files, start=1):
                if progress_callback:
                    progress_callback(number, total, up_file.name)

                out_name = f"{Path(up_file.name).stem}.xlsx"
                raw_bytes = up_file.getvalue()

                try:
                    if use_com and excel is not None:
                        local_in = work / f"in_{number}.xlsb"
                        local_out = work / f"out_{number}.xlsx"
                        workbook = None
                        try:
                            local_in.write_bytes(raw_bytes)
                            workbook = open_workbook(excel, local_in)
                            workbook.SaveAs(Filename=str(local_out), FileFormat=XLSX_FORMAT)
                            workbook.Close(SaveChanges=False)
                            workbook = None
                            converted_files.append((out_name, local_out.read_bytes()))
                        finally:
                            if workbook is not None:
                                try:
                                    workbook.Close(SaveChanges=False)
                                except Exception:
                                    pass
                            for temp_file in (local_in, local_out):
                                temp_file.unlink(missing_ok=True)
                    else:
                        # Cloud / Linux fallback
                        out_bytes = convert_via_python_engine(raw_bytes)
                        converted_files.append((out_name, out_bytes))

                except Exception as exc:
                    failed_files.append((up_file.name, describe(exc)))
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            del excel
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return converted_files, failed_files