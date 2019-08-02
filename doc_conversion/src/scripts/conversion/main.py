import os
import random


from win32com.client import Dispatch

from conversion.UltraSoundReport import UltraSound


def run_main():

    # DIR = "../../../data/20190722/"
    # DIR = "M:\\PycharmProjects\\doc_conversion\\data\\20190722\\"
    DIR = 'E:\\zjc\\'
    # Get all files in list
    word_documents = []
    excel_spreadsheets = []
    pdfs = []

    for version in os.listdir(DIR):
        if "/" in DIR:
            sub_dir = "{}{}/".format(DIR, version)
        else:
            sub_dir = "{}{}\\".format(DIR, version)
        for category in os.listdir(sub_dir):
            if "/" in DIR:
                ssub_dir = sub_dir + category + "/"
            else:
                ssub_dir = sub_dir + category + "\\"
            for item in os.listdir(ssub_dir):
                if "/" in DIR:
                    sssub_dir = ssub_dir + item + "/"
                else:
                    sssub_dir = ssub_dir + item + "\\"
                for file in os.listdir(sssub_dir):
                    if "~$" not in file:
                        if file.split(".")[-1].lower() in ["doc", "docx"]:
                            if file.split(".")[-2] in ["doc", "docx"]:
                                word_documents.append(sssub_dir + file[:-3])
                                continue
                            word_documents.append(sssub_dir + file)
                        elif file.split(".")[-1].lower() in ["xls", "xlsx"]:
                            excel_spreadsheets.append(sssub_dir + file)
                        elif file.split(".")[-1].lower() in ["pdf"]:
                            pdfs.append(sssub_dir + file)
    # 保证每次执行都一致
    random.seed(11)
    # 用于将一个列表中的元素打乱，即将列表中的元素随机排序
    random.shuffle(word_documents)
    word = Dispatch("Word.Application")

    good = 0
    err = 0
    # for file in word_documents:
    #     try:
    #         obj = UltraSound(file, word)
    #         good += 1
    #     except Exception:
    #         err += 1
    #     if (good + err) % 100 ==0:
    #         print("processed {}".format(good+err))
    # print(good, err, good/(good+err))
    # if obj.title != "颈动脉超声检查报告":
    # print(obj.title)

    # f = word.Documents.Open("M:\\PycharmProjects\\doc_conversion\\data\\20190722\\IV期\\Long_fu_survey_Carotid_ultrasound\\6201\\G6201305755报告单.doc")

    # for t in f.Tables:
    #     t.Cell(3,2)

    # for file in word_documents:
    #     try:
    #         obj = UltraSound(file, word)
    #         good += 1
    #     except Exception:
    #         err += 1
    #         print("\n\n==============================================={} err out of {} files===========================================\n\n".format(err, err+good))

    for file in word_documents:
        if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\王欣华G2202301132.doc":
            continue
        if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\郝守方G2202301624.doc":
            continue
        if file == "E:\\zjc\\IV期\\Long_fu_survey_Carotid_ultrasound\\2202\\王志伟G2202300829.doc":
            continue
        print(file)
        obj = UltraSound(file, word)


if __name__ == '__main__':
    run_main()
    pass
