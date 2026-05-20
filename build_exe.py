from pathlib import Path
import shutil

from PyInstaller.__main__ import run


APP_NAME = "DutyScheduler"

README_TEXT = """智能值班表排班系统 使用说明

一、启动
1. 双击 DutyScheduler.exe。
2. 程序会自动打开浏览器页面。如果浏览器没有自动打开，请稍等几秒后重新双击。
3. 使用过程中不要删除旁边的 _internal 文件夹。

二、使用
1. 在“数据导入”页面上传人员名单 Excel/CSV。
2. 批量导入所有人的课表 PDF，文件名需要包含姓名，例如：田宇(2025-2026-1)课表.pdf。
   支持三种导入方式：选择多个PDF文件、选择课表文件夹、填写本机文件夹路径。
3. 在“排班制作”页面选择周次并排班。
4. 下载导出的值班表 Excel。

三、数据保存与清空
导入的数据会自动保存在软件文件夹中的：
data\\app_state.json
如果你把整个 DutyScheduler 文件夹移动到别的位置，这份数据也会跟着移动。
在“数据导入”页面底部的“当前数据管理”中，可以删除某一份课表，或清空人员名单、课表、排班记录、全部数据。

四、人员名单表头
支持中文表头：姓名、专业班级/班级、部门、职务/职位、校区。
校区可以写“旗山北校区/旗山南校区”，程序会自动识别为北校区/南校区。

五、无法启动时
日志位置通常在：
%LOCALAPPDATA%\\DutyScheduler\\DutyScheduler.log
"""


def copy_release_files():
    release_dir = Path("dist") / APP_NAME
    release_dir.mkdir(parents=True, exist_ok=True)
    (release_dir / "使用说明.txt").write_text(README_TEXT, encoding="utf-8")

    sample_dir = release_dir / "示例文件"
    sample_dir.mkdir(exist_ok=True)
    for source in [
        Path("人员名单") / "人员名单模板.xlsx",
        Path("personnel.xlsx"),
    ]:
        if source.exists():
            shutil.copy2(source, sample_dir / source.name)


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
