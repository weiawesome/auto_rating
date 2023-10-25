import bs4
import pandas as pd


def get_id_score_pd(num=5, penalty=5, score=tuple([30, 50, 70, 80, 90, 100]),
                    result_file="./reference_files/index.html"):
    df = pd.DataFrame()
    exam_title = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
                  "U", "V", "W", "X", "Y", "Z"]

    result_file = open(result_file)
    file_content = result_file.read()
    result_file.close()

    html_content = bs4.BeautifulSoup(file_content, "html.parser")

    row_items = html_content.find_all("tr")

    content_title = []

    for item in row_items:

        title_items = item.find_all("th")
        for title in title_items:
            if title.text != "":
                content_title.append(str(title.text).strip())

        class_name = item.get("class")
        if class_name is None:
            continue
        if class_name is not None and "even" in class_name or "odd" in class_name:
            student_info = {}

            student_detail = item.find_all("td")
            for index, val in enumerate(student_detail):
                student_info[content_title[index]] = val.text
                if val.get("class") is not None and content_title[index] != "Name":
                    student_info[content_title[index]] = "yes" in val.get("class")

            student_info["Result"] = 0
            student_info["Fault"] = 0

            if str(student_info["Name"]).find("team") != -1 or student_info["Name"] == "TA":
                continue
            for n in range(num):
                if student_info[exam_title[n * 2]] or student_info[exam_title[n * 2 + 1]]:
                    student_info["Result"] += 1
                    if not (student_info[exam_title[n * 2]] and student_info[exam_title[n * 2 + 1]]):
                        student_info["Fault"] += 1

            student_info["Score"] = score[student_info["Result"]] - penalty * student_info["Fault"]
            df = pd.concat([df, pd.DataFrame(student_info, index=[0])], ignore_index=True)

    df["學號"] = df["Name"]
    df["分數"] = df["Score"]

    selected_columns = ["學號", "分數"]

    df = df[selected_columns]

    df = df.sort_values(by="學號", ascending=True)

    return df


def get_name_id_df(file_name="./reference_files/name_list.xlsx"):
    name_list_df = pd.read_excel(file_name)

    name_list_df["班級名稱"] = name_list_df["Unnamed: 1"]
    name_list_df["學號"] = name_list_df["Unnamed: 2"]
    name_list_df["姓名"] = name_list_df["Unnamed: 3"]

    name_list_df = name_list_df[["班級名稱", "學號", "姓名"]]

    name_list_df = name_list_df.dropna()
    name_list_df = name_list_df.iloc[1:]

    name_list_df = name_list_df.reset_index(drop=True)

    return name_list_df


def merge_result(name_list_df, id_score_df):
    df = pd.merge(id_score_df, name_list_df, on="學號", how="left")
    df = df[["班級名稱", "姓名", "學號", "分數"]]
    return df


def save_to_excel(df, result_file="./result/result.xlsx"):
    df.to_excel(result_file, index=False)
