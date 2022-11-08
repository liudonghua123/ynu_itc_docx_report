#%%
from os.path import dirname, join, realpath
from docxtpl import DocxTemplate
import fire
import jinja2
import pandas as pd
from dataclasses import dataclass
from datetime import datetime
from openpyxl import load_workbook

#%%
from config_logging import init_logging

logger = init_logging(join(dirname(realpath(__file__)), "main.log"))

#%%
# define a class to hold the data of each row
@dataclass
class Record:
    name: str
    gender: str
    id_num: str
    telphone: str
    company: str
    health_status: str
    car_num: str
    access_location: str
    access_date: str
    access_duration: str
    reason: str
    health_code_image_url: str
    travel_card_image_url: str
    nucleic_acid_testing_image_url: str
    health_pledge_image_url: str


# #%%
# # read the input data from the excel file using pandas.read_excel
# df = pd.read_excel(join(dirname(realpath(__file__)), "sample.xlsx"))
# logger.info(f"Read {len(df)} records from the excel file")
# logger.info(f"Head of the records:\n{df.head()}")
# # iterate the rows of the dataframe
# records: list[Record] = []
# for index, row in df.iterrows():
#     # create a Record object
#     record = Record(
#         name=row["姓名（必填）"],
#         gender=row["性别（必填）"],
#         id_num=row["身份证号码（必填）"],
#         telphone=row["手机号码（必填）"],
#         company=row["单位名称（必填）"],
#         health_status="健康",
#         car_num=row["车牌号码"],
#         access_location=row["到访地点（必填）"],
#         access_date=row["到访日期（必填）"],
#         access_duration=row["入校期限（必填）"],
#         reason=row["到访原因（必填）"],
#         health_code_image_url=row["云南省健康码（必填）"],
#         travel_card_image_url=row["行程卡截图（必填）"],
#         nucleic_acid_testing_image_url=row["核酸检测截图（必填）"],
#         health_pledge_image_url=row["《个人健康承诺书》（必填）"],
#     )
#     records.append(record)
# logger.info(f"Read {len(records)} records from the excel file.")

#%%
# utilities
def get_hyperlink_target(cell):
    result = None
    try:
        if cell != None and cell.hyperlink != None:
            result = cell.hyperlink.target
    except:
        result = None
    return result


#%%
def main(
    input_file_path: str = join(dirname(realpath(__file__)), "sample.xlsx"),
    output_file_path: str = join(dirname(realpath(__file__)), "generated_doc.docx"),
):
    # read the input data from the excel file using openpyxl
    workbook = load_workbook(filename=input_file_path)
    sheet = workbook.active

    # iterate the rows of the worksheet
    records: list[Record] = []
    for (
        submitter,
        submit_datetime,
        name,
        gender,
        id_num,
        telphone,
        company,
        car_num,
        access_location,
        access_date,
        access_duration,
        reason,
        health_code_image_url,
        health_code_image_url_detection,
        travel_card_image_url,
        travel_card_image_url_detection,
        nucleic_acid_testing_image_url,
        health_pledge_image_url,
    ) in sheet.iter_rows(min_row=2, max_col=18):
        # create a Record object
        record = Record(
            name=name.value,
            gender=gender.value,
            id_num=id_num.value,
            telphone=telphone.value,
            company=company.value,
            health_status="健康",
            car_num=car_num.value,
            access_location=access_location.value,
            access_date=access_date.value,
            access_duration=access_duration.value,
            reason=reason.value,
            health_code_image_url=get_hyperlink_target(health_code_image_url),
            travel_card_image_url=get_hyperlink_target(travel_card_image_url),
            nucleic_acid_testing_image_url=get_hyperlink_target(nucleic_acid_testing_image_url),
            health_pledge_image_url=get_hyperlink_target(health_pledge_image_url),
        )
        records.append(record)
    logger.info(f"Read {len(records)} records from the excel file.")

    #%%
    # generate the output docx file from the template with the data
    doc = DocxTemplate("template.docx")
    context = {"records": records}
    # pass some global utility functions/objects to the template
    jinja_env = jinja2.Environment()
    jinja_env.globals["len"] = len
    jinja_env.globals["datetime"] = datetime
    jinja_env.globals["enumerate"] = enumerate
    # render the template
    doc.render(context, jinja_env=jinja_env)
    # save the output docx file
    logger.info(f"Save the output docx file to {output_file_path}")
    doc.save(output_file_path)


if __name__ == "__main__":
    # Make Python Fire not use a pager when it prints a help text
    fire.core.Display = lambda lines, out: print(*lines, file=out)
    fire.Fire(main)
