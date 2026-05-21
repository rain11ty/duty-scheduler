from datetime import date
from pathlib import Path
import shutil

from PyInstaller.__main__ import run


APP_NAME = "DutyScheduler"


def copy_release_files():
    release_dir = Path("dist") / APP_NAME
    release_dir.mkdir(parents=True, exist_ok=True)

    readme_path = Path("README.md")
    if readme_path.exists():
        readme_text = readme_path.read_text(encoding="utf-8")
        shutil.copy2(readme_path, release_dir / "README.md")
        (release_dir / "使用说明.txt").write_text(readme_text, encoding="utf-8")

    sample_dir = release_dir / "示例文件"
    sample_dir.mkdir(exist_ok=True)
    for source in [
        Path("人员名单") / "人员名单模板.xlsx",
        Path("personnel.xlsx"),
    ]:
        if source.exists():
            shutil.copy2(source, sample_dir / source.name)


def make_release_zip():
    release_dir = Path("dist") / APP_NAME
    archive_base = Path("dist") / f"{APP_NAME}-Windows-{date.today():%Y%m%d}"
    archive_path = archive_base.with_suffix(".zip")
    if archive_path.exists():
        archive_path.unlink()
    return shutil.make_archive(
        str(archive_base),
        "zip",
        root_dir=release_dir.parent,
        base_dir=release_dir.name,
    )


if __name__ == "__main__":
    opts = [
        "run_main.py",
        f"--name={APP_NAME}",
        "--onedir",
        "--console",
        "--clean",
        "--noconfirm",
        "--add-data=duty_app.py;.",
        "--add-data=logic.py;.",
        "--collect-data=streamlit",
        "--collect-data=altair",
        "--copy-metadata=streamlit",
        "--copy-metadata=altair",
        "--copy-metadata=pandas",
        "--copy-metadata=numpy",
        "--copy-metadata=pyarrow",
        "--copy-metadata=pdfplumber",
        "--copy-metadata=openpyxl",
        "--copy-metadata=xlsxwriter",
        "--copy-metadata=watchdog",
        "--collect-submodules=streamlit",
        "--hidden-import=streamlit.web.cli",
        "--hidden-import=streamlit.runtime.scriptrunner.magic_funcs",
        "--hidden-import=numpy._core._exceptions",
        "--hidden-import=openpyxl",
        "--hidden-import=xlsxwriter",
        "--hidden-import=pdfplumber",
        "--exclude-module=torch",
        "--exclude-module=tensorflow",
        "--exclude-module=scipy",
        "--exclude-module=sklearn",
        "--exclude-module=statsmodels",
        "--exclude-module=matplotlib",
        "--exclude-module=pytest",
        "--exclude-module=IPython",
        "--exclude-module=jupyter",
        "--exclude-module=notebook",
        "--exclude-module=numba",
        "--exclude-module=plotly",
        "--exclude-module=bokeh",
        "--exclude-module=sqlalchemy",
        "--exclude-module=duckdb",
        "--exclude-module=polars",
        "--exclude-module=dask",
        "--exclude-module=pyspark",
        "--exclude-module=pandas.tests",
        "--exclude-module=pyarrow.tests",
        "--exclude-module=numpy.tests",
    ]

    run(opts)
    copy_release_files()
    archive = make_release_zip()
    print(f"Release archive created: {archive}")
