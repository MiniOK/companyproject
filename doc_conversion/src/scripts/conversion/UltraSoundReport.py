# -*- coding:utf-8 -*-
import os
import re
import warnings
from shutil import copyfile

import xlrd
from win32com.client import Dispatch


class UltraSound:

    def __init__(self, file_path, word):
        self.file_path = file_path
        self.title = "颈动脉超声检查报告"
        self.name = "Unknown"
        self.ID = "Unknown"
        self.gender = "Unknown"
        self.date = "Unknown"
        self.CCA_IMT_left = "Normal"
        self.plaques_count_left = "Normal"
        self.largest_plaque_width_left = "Normal"
        self.largest_plaque_depth_left = "Normal"
        self.plaque_shape_left = "Normal"
        self.plaque_is_ulcer_left = "Normal"
        self.plaque_texture_left = "Normal"
        self.DS_left = "Normal"
        self.location_left = "Normal"

        self.CCA_IMT_right = "Normal"
        self.plaques_count_right = "Normal"
        self.largest_plaque_width_right = "Normal"
        self.largest_plaque_depth_right = "Normal"
        self.plaque_shape_right = "Normal"
        self.plaque_is_ulcer_right = "Normal"
        self.plaque_texture_right = "Normal"
        self.DS_right = "Normal"
        self.location_right = "Normal"

        self.comments = ""
        self.doctor = "Unknown"

        self.valid = True
        extension = file_path.split(".")[-1]
        if extension.lower() in ["doc", "docx"]:
            self.file_type = "word"
            self.load_doc(word)
        elif extension.lower() in ["xls", "xlsx"]:
            self.file_type = "excel"
            self.load_xls()
        elif extension.lower() in ["pdf"]:
            self.file_type = "pdf"
        else:
            warn_msg = "Unknown file type with extension {} from {}. Please double check.".format(extension, file_path)
            warnings.warn(warn_msg)
            self.valid = False

    @staticmethod
    def fill(para, value):
        if value is None or value == "":
            return para
        else:
            return value

    def load_doc(self, word):
        # word = Dispatch("Word.Application.8")
        word.Visible = 0
        try:
            f = word.Documents.Open(self.file_path)
        except Exception:
            copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
            return 0
        content = f.Content.Text.replace("\t", "").replace("＝", "=").replace("\r", "").replace("\xa0", "").replace(
            "\x07", "").replace("官腔", "管腔").replace("\u3000", "").replace("\x00", "").replace("\x01",
                                                                                              "").replace(
            "\x15", "").replace("\x0c", "").replace("\x0e", "").replace("\x0c", "").replace("\x0b", "").replace(" ",
                                                                                                                "").replace(
            ":", "").replace("：", "").replace("端", "段")
        if content[:1] != "广西医科大学附属武鸣医院":
            print("content", content)
            if len(content) < 10:
                copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                return 0
            # self.title = self.fill(self.title, re.search("南宁市武鸣区人民医院(.*?)姓名：", content).group(1).replace("\r", ""))
            try:
                self.name = self.fill(self.name, re.search("姓名(.*?)性别", content).group(1).replace(" ", ""))
                if len(self.name) > 5:
                    raise ValueError
            except Exception:
                self.name = self.fill(self.name, re.search("姓名(.*?)(受检者ID|受检查ID)", content).group(1).replace(" ", ""))
            # print(self.name)
            try:
                self.ID = self.fill(self.ID, re.search("患者ID(.*?)左侧CCA-IMT", content).group(1))
            except Exception:
                try:
                    self.ID = self.fill(self.ID, re.search("受检者id|受检者ID(.*?)性别", content).group(1))
                except Exception:
                    copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0
            # print(self.ID)
            try:
                self.gender = self.fill(self.gender, re.search("性别(.*?)年龄", content).group(1))
            except Exception:
                self.gender = self.fill(self.gender, re.search("性别(.*?)检查日期", content).group(1))
            # print(self.gender)
            try:
                self.date = self.fill(self.date, re.search("检查日期(.*?)打印日期", content).group(1))
            except Exception:
                self.date = self.fill(self.date, re.search("检查日期(.*?)(左侧|右侧|LCCA-IMT|2D及M型)", content).group(1))
            # print(self.date)
            if "LCCA-IMT" in content:
                if content.index("LCCA-IMT") < content.index("RCCA-IMT"):
                    left = re.search("LCCA-IMT(.*?)RCCA-IMT", content).group(1)
                    right = re.search("RCCA-IMT(.*?)超声印象", content).group(1)
                else:
                    right = re.search("RCCA-IMT(.*?)LCCA-IMT", content).group(1)
                    left = re.search("LCCA-IMT(.*?)超声印象", content).group(1)
            else:
                try:
                    if content.index("左侧") < content.index("右侧"):
                        left = re.search("左侧(.*?)右侧", content).group(1)

                        right = re.search("右侧(.*?)(超声印象|检查医生|报告医生)", content).group(1)     # 添加 报告医生
                    else:
                        right = re.search("右侧(.*?)左侧", content).group(1)
                        left = re.search("左侧(.*?)(超声印象|检查提示)", content).group(1)
                except ValueError:
                    copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0
            print("left", left)
            try:
                self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], re.search("近段(.*?)中段", left).group(1)),
                                     self.fill(self.CCA_IMT_left[1], re.search("中段(.*?)远段", left).group(1)),
                                     self.fill(self.CCA_IMT_left[2], re.search("远段(.*?)(数量|斑块)", left).group(1)))

            except Exception:
                try:
                    self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], re.search("近段(.*?)mm", left).group(1)),
                                         self.fill(self.CCA_IMT_left[1], re.search("中段(.*?)mm", left).group(1)),
                                         self.fill(self.CCA_IMT_left[2], re.search("远段(.*?)mm", left).group(1)))
                except Exception:
                    copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
                    return 0
            # print(self.CCA_IMT_left)
            try:
                self.plaques_count_left = self.fill(self.plaques_count_left,
                                                    re.search("数量（1=单发，2=多发）\[(.*?)\]", left).group(1))
            except Exception:
                try:
                    self.plaques_count_left = self.fill(self.plaques_count_left,
                                                        re.search("数量1.无2.单发3.多发(.*?)最大者", left).group(1))
                except Exception:
                    try:
                        self.plaques_count_left = self.fill(self.plaques_count_left,
                                                            re.search("(数量（1=单发，2=多发）|数量（1＝单发，2＝多发）)(.*?)(最大者长度|最长者长度)",
                                                                      left).group(1))
                    except Exception:
                        self.plaques_count_left = self.fill(self.plaques_count_left,
                                                            re.search("(数量（1=单发，2=多发）|数量（1＝单发，2＝多发）)(.*?)(最大者长度|最大者厚度)",
                                                                      left).group(1))
            # print(self.plaques_count_left)
            try:
                self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
                                                           re.search("最大者长度(.*?)mm", left).group(1))
            except Exception:
                self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
                                                           re.search("(大者长度|最大者长|最大者长度|最长者长度)(.*?)(最大者厚度|最大厚度)", left).group(1)) # 添加大者长度
            # print(self.largest_plaque_width_left)
            try:
                self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                           re.search("最大者厚度|大者厚度(.*?)mm。", left).group(1))
            except Exception:
                try:
                    self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                               re.search("最大者厚度(.*?)回声", left).group(1))
                except Exception:
                    self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
                                                               re.search("最大厚度(.*?)形态", left).group(1))
            # print(self.largest_plaque_depth_left)
            try:
                self.plaque_shape_left = self.fill(self.plaque_shape_left,
                                                   re.search("形态（1=规则型，2=不规则型）[(.*?)]", left).group(1))
            except Exception:
                try:
                    self.plaque_shape_left = self.fill(self.plaque_shape_left,
                                                       re.search("形态1.规则型2.不规则型(.*?)有无溃疡斑块", left).group(1))
                except Exception:
                    self.plaque_shape_left = self.fill(self.plaque_shape_left,
                                                       re.search(
                                                           "(形态（1=规则型，2=不规则型|形态（1=规则型，2=不规则型）|（1=规则型，2=不规则）|形态（1=规则型，2=不规则型。）)(.*?)是否溃疡型",
                                                           left).group(1))
            # print(self.plaque_shape_left)
            try:
                self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left,
                                                      re.search("是否溃疡型（0=否，1=是）[(.*?)]", left).group(1))
            except Exception:
                try:
                    self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left,
                                                          re.search("有无溃疡斑块1.无2.有(.*?)狭窄程度", left).group(1))
                except Exception:
                    self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left,
                                                          re.search("(是否溃疡型（0=否，1=是）|是否溃疡型（0＝否，1＝是）)(.*?)质地",
                                                                    left).group(1))
            # print(self.plaque_is_ulcer_left)
            try:
                self.plaque_texture_left = self.fill(self.plaque_texture_left,
                                                     re.search("质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质A3|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声2，B=不均质）(.*?)管腔直径狭窄率%",
                                                     # re.search("质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）[(.*?)]", # 此处有问题 先注掉
                                                               left).group(1))
            except Exception:
                try:
                    self.plaque_texture_left = self.fill(self.plaque_texture_left,
                                                         re.search("1.强回声2.中等回声3.低回声4.不均匀回声(.*?)形态1.规则型", left).group(
                                                             1))
                except Exception:
                    self.plaque_texture_left = self.fill(self.plaque_texture_left,
                                                         re.search(
                                                             "(质地（A1=均质低回声，A2=均质等1回声，A3=均质强回声，B=不均质）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）\[)(.*?)(官腔直径狭窄率|管腔直径狭窄率|]。管腔直径狭窄率)",
                                                             left).group(
                                                             1))
            # print(self.plaque_texture_left)
            try:
                self.DS_left = self.fill(self.DS_left, re.search("管腔直径狭窄率(.*?)狭窄部位", left).group(1))
            except Exception:
                try:
                    self.DS_left = self.fill(self.DS_left, re.search("狭窄程度或闭塞部位(.*?)检查结果", left).group(1))
                except Exception:
                    self.DS_left = self.fill(self.DS_left,
                                             re.search("(管腔直径狭窄率|官腔直径狭窄率)(.*?)(狭窄部位)", left).group(1))
            # print(self.DS_left)
            try:
                self.location_left = self.fill(self.location_left, re.search("狭窄部位(.*?)", left).group(1))
            except Exception:
                self.location_left = self.fill(self.location_left, re.search("狭窄程度或闭塞部位(.*?)检查结果", left).group(1))
            print(self.location_left)









            # print("right", right)
            # try:
            #     self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("近段(.*?)中段", right).group(1)),
            #                           self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)远段", right).group(1)),
            #                           self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)(数量|斑块)", right).group(1)))
            # except Exception as e:
            #     try:
            #         self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], re.search("近段(.*?)mm", right).group(1)),
            #                               self.fill(self.CCA_IMT_right[1], re.search("中段(.*?)mm", right).group(1)),
            #                               self.fill(self.CCA_IMT_right[2], re.search("远段(.*?)mm", right).group(1)))
            #     except Exception:
            #         copyfile(self.file_path, "../output/err/{}".format(self.file_path.split("\\")[-1]))
            #         return 0
            # # print(self.CCA_IMT_right)
            # try:
            #     self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                          re.search("数量（1=单发，2=多发）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                              re.search("数量1.无2.单发3.多发(.*?)(最大者长度|最长者长度)", right).group(1))
            #     except Exception:
            #         self.plaques_count_right = self.fill(self.plaques_count_right,
            #                                              re.search("数量（1=单发，2=多发）(.*?)(最大者长度|最长者长度)", right).group(1))
            # # print(self.plaques_count_right)
            # try:
            #     self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
            #                                                 re.search("(最大者长度|最长者长度)(.*?)mm", right).group(1))
            # except Exception:
            #     try:
            #         self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
            #                                                     re.search("(最大者长度|最长者长度)(.*?)(最大者厚度|最大厚度|最大着厚度)",
            #                                                               right).group(1))
            #     except Exception:
            #         raise Exception
            # # print(self.largest_plaque_width_right)
            # try:
            #     self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
            #                                                 re.search("最大者厚度(.*?)mm。", right).group(1))
            # except Exception:
            #     try:
            #         self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
            #                                                     re.search("最大者厚度(.*?)回声", right).group(1))
            #     except Exception:
            #         # try:
            #         regex = re.compile("(最大着厚度|最大者厚度|最大厚度)(.*?)形态")
            #         self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
            #                                                    regex.search(right).group(1))
            #             # 修改 NoneType 异常
            #         # except Exception:
            #         #     warn_msg = "Warnings come form {} ".format(self.file_path)
            #         #
            #         #     warnings.warn(warn_msg)
            #         #     import time
            #         #
            #         #     time.sleep(60)
            #         #     print('正在等待。。。')
            # # print(self.largest_plaque_depth_right)
            # try:
            #     self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                         re.search("形态（1=规则型，2=不规则型）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                             re.search("形态1.规则型2.不规则型(.*?)有无溃疡斑块", right).group(1))
            #     except Exception:
            #         self.plaque_shape_right = self.fill(self.plaque_shape_right,
            #                                             re.search(
            #                                                 "(形态（1=规则型，2=不规则型|形态（1=规则型，2=不规则型）|形态（1=规则型，2=不规则）|形态（1=规则型，2=不规则型。）)(.*?)是否溃疡",
            #                                                 right).group(1))
            # # print(self.plaque_shape_right)
            # try:
            #     self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                            re.search("是否溃疡型（0=否，1=是）[(.*?)]", right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                                re.search("有无溃疡斑块1.无2.有(.*?)狭窄程度", right).group(1))
            #     except Exception:
            #         self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right,
            #                                                re.search("(是否溃疡型（0=否，1=是）|是否溃疡（0=否，1=是）)(.*?)质地",
            #                                                          right).group(1))
            # # print(self.plaque_is_ulcer_right)
            # try:
            #     self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                           re.search("质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）[(.*?)]",
            #                                                     right).group(1))
            # except Exception:
            #     try:
            #         self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                               re.search("1.强回声2.中等回声3.低回声4.不均匀回声(.*?)形态1.规则型", right).group(
            #                                                   1))
            #     except Exception:
            #         self.plaque_texture_right = self.fill(self.plaque_texture_right,
            #                                               re.search(
            #                                                   "(质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均匀）|质地（A1均质低回声，A2均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3=均质强回声，B=不均质）|质地（A1=均质低回声，A2=均质等回声，A3均质强回声，B=不均质）\[)(.*?)(管腔直径狭窄率|]。管腔直径狭窄率|狭窄部位)",
            #                                                   right).group(
            #                                                   1))
            #
            # # print(self.plaque_texture_right)
            # try:
            #     self.DS_right = self.fill(self.DS_right, re.search("管腔直径狭窄率(.*?)(%狭窄部位|狭窄部位)", right).group(1))
            # except Exception:
            #     try:
            #         self.DS_right = self.fill(self.DS_right, re.search("狭窄程度或闭塞部位(.*?)(检查结果|左侧)", right).group(1))
            #     except Exception:
            #         try:
            #             self.DS_right = self.fill(self.DS_right, re.search("管腔直径狭窄率(.*?)狭窄部位", right).group(1))
            #         except Exception:
            #             self.DS_right = self.fill(self.DS_right, re.search("(狭窄程度或闭塞部位|狭窄部位)(.*?)", right).group(1))
            # # print(self.DS_right)
            # try:
            #     self.location_right = self.fill(self.location_right, re.search("狭窄部位(.*?)", right).group(1))
            # except Exception:
            #     try:
            #         self.location_right = self.fill(self.location_right,
            #                                         re.search("狭窄程度或闭塞部位(.*?)检查结果", right).group(1))
            #     except Exception:
            #         self.location_right = self.fill(self.location_right,
            #                                         re.search("狭窄程度或闭塞部位(.*?)左侧", content).group(1))
            # # print(self.location_right)
            #
            # try:
            #     self.comments = self.fill(self.comments, re.search("狭窄部位|超声印象(.*?)", content).group(1)) # 添加 空串 ''
            # except Exception as e:
            #     print(e)
            #     self.comments = self.fill(self.comments, re.search("检查提示(.*?)检查医生", content).group(1))
            # for t in f.Tables:
            #     try:
            #         self.doctor = self.fill(self.doctor, t.Cell(20, 7).Range.Text)
            #     except Exception:
            #         # self.doctor = re.search(content, "报告医生：(.*?)").group(1)
            #         try:
            #             self.doctor = re.search("(报告医生|报告医生:|报告医师:)(.*?)\r", f.Content.Text).group(1)
            #         except Exception:
            #             try:
            #                 self.doctor = re.search("(报告医生 |报告医师 |检查医生)(.*?)\r", f.Content.Text).group(1)
            #             except Exception:
            #                 try:
            #                     self.doctor = re.search("\r\x07\r\x07\r(报告医生：|报告医师：)(.*?)报告机构", f.Content.Text).group(
            #                         1).replace(" ", "")
            #                 except Exception:
            #                     self.doctor = re.search("(报告医生：|报告医师：)(.*?)\r", f.Content.Text).group(
            #                         1).replace(" ", "")
            #     break
            # print(self.doctor)
            # f.Close()







        # elif content[:6] == "濮阳市中医院":
        #     content = content.replace("\t", "").replace("\r", "").replace("\xa0", "").replace("\x07", "").replace(
        #         "\x01", "").replace("\x15", "").replace("\x0c", "").replace("\x0e", "").replace("\x0c", "")
        #     self.title = self.fill(self.title, re.search("濮阳市中医院(.*?)姓名", content).group(1).replace("\r", ""))
        #     self.name = self.fill(self.name, re.search("姓名：(.*?)受检者ID", content).group(1).replace(" ", ""))
        #     self.ID = self.fill(self.ID, re.search("受检者ID：(.*?)别", content).group(1))
        #     self.gender = self.fill(self.gender, re.search("性别：(.*?)检查日期", content).group(1))
        #     self.date = self.fill(self.date, re.search("检查日期：(.*?)左侧", content).group(1))
        #     for t in f.Tables:
        #         self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], t.Cell(3, 2).Range.Text.split("\t")[0]),
        #                              self.fill(self.CCA_IMT_left[1], t.Cell(3, 3).Range.Text.split("\t")[0]),
        #                              self.fill(self.CCA_IMT_left[2], t.Cell(3, 4).Range.Text))
        #         left = re.search("左侧(.*?)右侧", content).group(1)
        #         self.plaques_count_left = self.fill(self.plaques_count_left,
        #                                             re.search("数量（1=单发，2=多发）(.*?)最大者长度", left).group(1))
        #         self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
        #                                                    re.search("最大者长度：(.*?)", left).group(1))
        #         self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
        #                                                    re.search("最大者厚度：(.*?)形态", left).group(1))
        #         self.plaque_shape_left = self.fill(self.plaque_shape_left, t.Cell(6, 2).Range.Text[2:3])
        #         self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left, t.Cell(6, 3).Range.Text)
        #         self.plaque_texture_left = self.fill(self.plaque_texture_left,
        #                                              re.search("A2=均质等回声，A3=均质强回声，B=不均质）：(.*?)管腔直径狭窄率", left).group(1))
        #         self.DS_left = self.fill(self.DS_left, t.Cell(8, 2).Range.Text)
        #         self.location_left = self.fill(self.location_left, t.Cell(8, 4).Range.Text)
        #
        #         right = re.search("右侧(.*?)报告医师", content).group(1)
        #         self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], t.Cell(11, 2).Range.Text),
        #                               self.fill(self.CCA_IMT_right[1], t.Cell(11, 6).Range.Text),
        #                               self.fill(self.CCA_IMT_right[2], t.Cell(11, 8).Range.Text))
        #         self.plaques_count_left = self.fill(self.plaques_count_left,
        #                                             re.search("数量（1=单发，2=多发）(.*?)最大者长度", right).group(1))
        #         self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
        #                                                    re.search("最大者长度：(.*?)形态", right).group(1))
        #         self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
        #                                                    re.search("最大者厚度：(.*?)形态", right).group(1))
        #         self.plaque_shape_right = self.fill(self.plaque_shape_right, t.Cell(14, 2).Range.Text[2:3])
        #         self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right, t.Cell(14, 3).Range.Text)
        #         self.plaque_texture_right = self.fill(self.plaque_texture_right,
        #                                               re.search("A3=均质强回声，B=不均质）：(.*?)管腔直径狭窄率", right).group(1))
        #         self.DS_right = self.fill(self.DS_right, t.Cell(16, 2).Range.Text)
        #         self.location_right = self.fill(self.location_right, t.Cell(16, 4).Range.Text)
        #
        #         self.comments = self.fill(self.comments, re.search("超声印象：\r(.*?)\r", f.Content.Text).group(1))
        #         try:
        #             self.doctor = self.fill(self.doctor, t.Cell(20, 7).Range.Text)
        #         except Exception:
        #             self.doctor = re.search("报告医师： (.*?)\r", f.Content.Text).group(1)
        #             # print(self.doctor)
        #         f.Close()
        #         break
        # else:
        #     for t in f.Tables:
        #         if t.Cell(1, 1).Range.Text.replace("\r\x07", "") == "颈动脉超声检查报告":
        #             self.title = self.fill(self.title, t.Cell(1, 1).Range.Text.replace("\r\x07", ""))
        #             self.name = self.fill(self.name, t.Cell(2, 2).Range.Text)
        #             self.ID = self.fill(self.ID, t.Cell(2, 4).Range.Text)
        #             self.gender = self.fill(self.gender, t.Cell(2, 6).Range.Text)
        #             self.date = self.fill(self.date, t.Cell(2, 8).Range.Text)
        #             self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], t.Cell(5, 2).Range.Text),
        #                                  self.fill(self.CCA_IMT_left[1], t.Cell(5, 4).Range.Text),
        #                                  self.fill(self.CCA_IMT_left[2], t.Cell(5, 6).Range.Text))
        #             self.plaques_count_left = self.fill(self.plaques_count_left, t.Cell(7, 2).Range.Text)
        #             self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left, t.Cell(7, 4).Range.Text)
        #             self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left, t.Cell(7, 6).Range.Text)
        #             self.plaque_shape_left = self.fill(self.plaque_shape_left, t.Cell(8, 2).Range.Text)
        #             self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left, t.Cell(8, 4).Range.Text)
        #             try:
        #                 self.plaque_texture_left = self.fill(self.plaque_texture_left, t.Cell(9, 2).Range.Text)
        #             except Exception:
        #                 pass
        #             self.DS_left = self.fill(self.DS_left, t.Cell(10, 2).Range.Text)
        #             self.location_left = self.fill(self.location_left, t.Cell(10, 4).Range.Text)
        #
        #             self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], t.Cell(13, 2).Range.Text),
        #                                   self.fill(self.CCA_IMT_right[1], t.Cell(13, 4).Range.Text),
        #                                   self.fill(self.CCA_IMT_right[2], t.Cell(13, 6).Range.Text))
        #             self.plaques_count_right = self.fill(self.plaques_count_right, t.Cell(15, 2).Range.Text)
        #             self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
        #                                                         t.Cell(15, 4).Range.Text)
        #             self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
        #                                                         t.Cell(15, 6).Range.Text)
        #             self.plaque_shape_right = self.fill(self.plaque_shape_right, t.Cell(16, 2).Range.Text)
        #             self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right, t.Cell(16, 4).Range.Text)
        #             try:
        #                 self.plaque_texture_right = self.fill(self.plaque_texture_right, t.Cell(17, 2).Range.Text)
        #             except Exception:
        #                 pass
        #             self.DS_right = self.fill(self.DS_right, t.Cell(18, 2).Range.Text)
        #             self.location_right = self.fill(self.location_right, t.Cell(18, 4).Range.Text)
        #
        #             self.comments = self.fill(self.comments, t.Cell(19, 1).Range.Text)
        #             try:
        #                 self.doctor = self.fill(self.doctor, t.Cell(20, 7).Range.Text)
        #             except Exception:
        #                 # self.doctor = re.search(content, "报告医生：(.*?)").group(1)
        #                 try:
        #                     self.doctor = re.search("报告医生:(.*?)\r", content).group(1)
        #                 except Exception:
        #                     try:
        #                         self.doctor = re.search("报告医生 (.*?)\r", content).group(1)
        #                     except Exception:
        #                         try:
        #                             self.doctor = re.search("\r\x07\r\x07\r报告医生：(.*?)报告机构", content).group(1).replace(
        #                                 " ", "")
        #                         except Exception:
        #                             self.doctor = re.search("报告医生：(.*?)\r", content).group(
        #                                 1).replace(" ", "")
        #                 # print(self.doctor)
        #             f.Close()
        #             break
        #         else:
        #             self.title = "颈动脉超声检查报告"
        #             self.name = self.fill(self.name, t.Cell(1, 2).Range.Text)
        #             self.ID = self.fill(self.ID, t.Cell(1, 4).Range.Text)
        #             self.gender = self.fill(self.gender, t.Cell(1, 6).Range.Text)
        #             self.date = self.fill(self.date, t.Cell(1, 8).Range.Text)
        #             self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], t.Cell(4, 2).Range.Text),
        #                                  self.fill(self.CCA_IMT_left[1], t.Cell(4, 4).Range.Text),
        #                                  self.fill(self.CCA_IMT_left[2], t.Cell(4, 6).Range.Text))
        #             self.plaques_count_left = self.fill(self.plaques_count_left, t.Cell(6, 2).Range.Text).replace("\r",
        #                                                                                                           "").replace(
        #                 "\x07", "")
        #             if int(self.plaques_count_left) > 1:
        #                 self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left,
        #                                                            t.Cell(6, 4).Range.Text)
        #                 self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left,
        #                                                            t.Cell(6, 6).Range.Text)
        #             self.plaque_shape_left = self.fill(self.plaque_shape_left, t.Cell(7, 2).Range.Text)
        #             self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left, t.Cell(7, 4).Range.Text)
        #             self.plaque_texture_left = self.fill(self.plaque_texture_left, t.Cell(8, 2).Range.Text)
        #             self.DS_left = self.fill(self.DS_left, t.Cell(9, 2).Range.Text)
        #             self.location_left = self.fill(self.location_left, t.Cell(9, 4).Range.Text)
        #             try:
        #                 self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], t.Cell(12, 2).Range.Text),
        #                                       self.fill(self.CCA_IMT_right[1], t.Cell(12, 4).Range.Text),
        #                                       self.fill(self.CCA_IMT_right[2], t.Cell(12, 6).Range.Text))
        #             except Exception:
        #                 self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], t.Cell(13, 2).Range.Text),
        #                                       self.fill(self.CCA_IMT_right[1], t.Cell(13, 4).Range.Text),
        #                                       self.fill(self.CCA_IMT_right[2], t.Cell(13, 6).Range.Text))
        #             try:
        #                 self.plaques_count_right = self.fill(self.plaques_count_right, t.Cell(14, 2).Range.Text)
        #             except Exception:
        #                 self.plaques_count_right = self.fill(self.plaques_count_right, t.Cell(15, 2).Range.Text)
        #             self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right,
        #                                                         t.Cell(14, 4).Range.Text)
        #             self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right,
        #                                                         t.Cell(14, 6).Range.Text)
        #             self.plaque_shape_right = self.fill(self.plaque_shape_right, t.Cell(15, 2).Range.Text)
        #             self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right, t.Cell(15, 4).Range.Text)
        #             self.plaque_texture_right = self.fill(self.plaque_texture_right, t.Cell(16, 2).Range.Text)
        #             self.DS_right = self.fill(self.DS_right, t.Cell(17, 2).Range.Text)
        #             self.location_right = self.fill(self.location_right, t.Cell(17, 4).Range.Text)
        #
        #             self.comments = self.fill(self.comments, t.Cell(18, 1).Range.Text)
        #             try:
        #                 self.doctor = self.fill(self.doctor, t.Cell(20, 7).Range.Text)
        #             except Exception:
        #                 # self.doctor = re.search(content, "报告医生：(.*?)").group(1)
        #                 self.doctor = re.search(content, "报告医生：(.*?)\r").group(1)
        #                 # print(self.doctor)
        #             f.Close()
        #             break

    def load_xls(self):
        # open the work book
        wb = xlrd.open_workbook(self.file_path)

        # open the sheet you want to read its cells
        sht = wb.sheet_by_index(0)
        k = 0
        self.title = self.fill(self.title, sht.cell(0, 0).value)
        if "超声心动图" in self.title:
            return 0
        if sht.cell(2, 0).value == "":
            k += 1
        self.name = self.fill(self.name, sht.cell(1, 1).value)
        self.ID = self.fill(self.ID, sht.cell(1, 3).value)
        self.gender = self.fill(self.gender, sht.cell(1, 5).value)
        self.date = self.fill(self.date, sht.cell(1, 7).value)
        self.CCA_IMT_left = (self.fill(self.CCA_IMT_left[0], sht.cell(4 + k, 1).value),
                             self.fill(self.CCA_IMT_left[1], sht.cell(4 + k, 3).value),
                             self.fill(self.CCA_IMT_left[2], sht.cell(4 + k, 5).value))
        self.plaques_count_left = self.fill(self.plaques_count_left, sht.cell(6 + k, 3).value)
        self.largest_plaque_width_left = self.fill(self.largest_plaque_width_left, sht.cell(6 + k, 5).value)
        self.largest_plaque_depth_left = self.fill(self.largest_plaque_depth_left, sht.cell(6 + k, 7).value)
        self.plaque_shape_left = self.fill(self.plaque_shape_left, sht.cell(7 + k, 3).value)
        self.plaque_is_ulcer_left = self.fill(self.plaque_is_ulcer_left, sht.cell(7 + k, 7).value)
        self.plaque_texture_left = self.fill(self.plaque_texture_left, sht.cell(8 + k, 7).value)
        self.DS_left = self.fill(self.DS_left, sht.cell(9 + k, 3).value)
        self.location_left = self.fill(self.location_left, sht.cell(9 + k, 5).value)

        if sht.cell(11, 0).value == "":
            k += 1
        self.CCA_IMT_right = (self.fill(self.CCA_IMT_right[0], sht.cell(12 + k, 1).value),
                              self.fill(self.CCA_IMT_right[1], sht.cell(12 + k, 3).value),
                              self.fill(self.CCA_IMT_right[2], sht.cell(12 + k, 5).value))
        self.plaques_count_right = self.fill(self.plaques_count_right, sht.cell(14 + k, 3).value)
        self.largest_plaque_width_right = self.fill(self.largest_plaque_width_right, sht.cell(14 + k, 5).value)
        self.largest_plaque_depth_right = self.fill(self.largest_plaque_depth_right, sht.cell(14 + k, 7).value)
        self.plaque_shape_right = self.fill(self.plaque_shape_right, sht.cell(15 + k, 3).value)
        self.plaque_is_ulcer_right = self.fill(self.plaque_is_ulcer_right, sht.cell(15 + k, 7).value)
        self.plaque_texture_right = self.fill(self.plaque_texture_right, sht.cell(16 + k, 7).value)
        self.DS_right = self.fill(self.DS_right, sht.cell(17 + k, 3).value)
        self.location_right = self.fill(self.location_right, sht.cell(17 + k, 5).value)

        if sht.cell(11, 0).value == "":
            k += 1
        # print(k)
        self.comments = self.fill(self.comments, sht.cell(19 + k, 0).value)
        self.doctor = self.fill(self.doctor, sht.cell(21 + k, 6).value)


if __name__ == "__main__":
    word = Dispatch("Word.Application")
    test = UltraSound(
        # r"E:\zjc\IV期\High_risk_Carotid_ultrasound\6306\104.doc",
        # r'E:\zjc\IV期\Long_fu_survey_Carotid_ultrasound\2302\G2302303532.doc',
        # r"E:\zjc\IV期\Long_fu_survey_Carotid_ultrasound\4505\G450516811.docx",
        # r"E:\zjc\IV期\\Long_fu_survey_Carotid_ultrasound\\5203\\颈动脉超声G520307411.doc",
        # r'E:\zjc\IV期\Long_fu_survey_Carotid_ultrasound\2202\王欣华G2202301132.doc',
        r'E:\zjc\IV期\Long_fu_survey_Carotid_ultrasound\2201\G220181738.doc',

        word)

