from operation_component import *
from PyQt5.QtWidgets import QWidget, QPushButton
from PyQt5.QtCore import Qt
import pyautogui
import pandas as pd
import  xlrd
import numpy
import os


class MainWindow(QWidget):

    def __init__(self, height, width, location, title):
        super().__init__()
        self._operation_list = []
        self._current_operation_index = 0
        self._current_data_position = 0
        self._start_btn = None
        self._is_started = False

        # 初始化数据
        input_data_path = r"C:\Users\TimCat\Desktop\jin_nian.xlsx"
        xls_data = pd.DataFrame(pd.read_excel(input_data_path))
        self._ids = list(xls_data['身份证号'])
        self._amounts = list(xls_data['合同余额'])
        self._names = list(xls_data['学生姓名'])
        self.init_current_position()
        self.add_all_operation()
        self.initUI(height, width, location, title)

    def initUI(self, height, width, location, title):
        self.setGeometry(location.x, location.y, width, height)
        self.setWindowTitle(title)
        btns_start_x = 30
        btns_start_y = 30
        y_offset = 30
        btn_start_title = "开始"
        btn_start_location = Point(btns_start_x, btns_start_y)
        btn_start = QPushButton(btn_start_title, self)
        btn_start.resize(btn_start.sizeHint())
        btn_start.move(btn_start_location.x, btn_start_location.y)
        btn_start.clicked.connect(self.start_operation)
        self._start_btn = btn_start

        btn_last_step_title = "上一步"
        btn_last_step_location = Point(btns_start_x, btns_start_y + y_offset)
        btn_last_step = QPushButton(btn_last_step_title, self)
        btn_last_step.resize(btn_last_step.sizeHint())
        btn_last_step.move(btn_last_step_location.x,  btn_last_step_location.y)
        btn_last_step.clicked.connect(self.last_operation)

        btn_next_step_title = "下一步"
        btn_next_step_location = Point(btns_start_x, btns_start_y + y_offset*2)
        btn_next_step = QPushButton(btn_next_step_title, self)
        btn_next_step.resize(btn_next_step .sizeHint())
        btn_next_step.move(btn_next_step_location.x, btn_next_step_location.y)
        btn_next_step.clicked.connect(self.next_operation)

        btn_success_title = "成功"
        btn_success_location = Point(btns_start_x, btns_start_y + y_offset * 3)
        btn_success = QPushButton(btn_success_title, self)
        btn_success.resize(btn_success.sizeHint())
        btn_success.move(btn_success_location.x, btn_success_location.y)
        btn_success.clicked.connect(self.operation_success)

        # 窗体总在最前端
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.show()

    def operation_success(self):
        if self._current_operation_index < len(self._operation_list) \
                and 0 < self._current_operation_index:
            print("请先完成操作")
            return
        self._current_data_position += 1
        print("数据位置： " + str(self._current_data_position))
        self._current_operation_index = 0
        self._start_btn.setText("开始")
        self._is_started = False
        self._operation_list = []
        self.add_all_operation()

    def init_current_position(self):
        store_file_path = "current_location.txt"
        if os.path.exists(store_file_path):
            with open(store_file_path, encoding='utf-8') as f:
                name = f.read().strip('\n')
                print("当前人的姓名" + name)
                if name in self._names:
                    self._current_operation_index = self._names.index(name)
                    print("所在位置：" + str(self._current_operation_index))
                else:
                    print('初始化失败')
                    raise KeyboardInterrupt

    # 关闭之前先保存数据
    def closeEvent(self, event):
        self.save_data()

    def start_operation(self):
        restart_title = "重新开始"
        start_title = "开始"
        self._current_operation_index = 0
        if self._is_started:
            self._start_btn.setText(start_title)
            self._is_started = False
        else:
            self._start_btn.setText(restart_title)
            self.one_operation(self._operation_list[0])
            self._current_operation_index += 1
            self._is_started = True

    def next_operation(self):
        if self._current_operation_index < len(self._operation_list) \
                and 0 < self._current_operation_index:
            index = self._current_operation_index
            self._current_operation_index += 1
            self.one_operation(self._operation_list[index])
        if self._current_operation_index == len(self._operation_list):
            print("执行完毕")

        elif self._current_operation_index == 0:
            print("请先开始")
        else:
            print("程序发生不可控错误")

    def last_operation(self):
        if 0 < self._current_operation_index < len(self._operation_list):
            self._current_operation_index -= 1
        else:
            print("没有上一步")

    def one_operation(self, operation):
        # 执行一个操作
        try:
            position = pyautogui.position()
            operation.operate()
            pyautogui.moveTo(position[0], position[1])
        except pyautogui.FailSafeException:
            print("鼠标移动到左上角，程序已停")
        except KeyboardInterrupt:
            print("按下终止键，程序已暂停")
        finally:
            self.save_data()

    def save_data(self):
        with open("current_location.txt", "w", encoding='utf-8') as f:
            f.write(self._names[self._current_data_position])

    def add_all_operation(self):
        amount = str(int(self._amounts[self._current_data_position])//10)
        id = str(self._ids[self._current_data_position])
        operation_list = self._operation_list
        # 新增按钮的位置
        btn_new_apply_location = Point(point_tuple=(56, 953))
        # 学生名称旁边的按钮的位置
        btn_pop_search_student_window_location = Point(point_tuple=(878, 358))
        # id输入框的位置
        inbx_id_location = Point(point_tuple=(761, 522))
        # 查询按钮位置
        btn_search_student_location = Point(point_tuple=(1164, 522))
        # 选择所查出来的学生按钮位置
        btn_specified_student_location = Point(point_tuple=(488, 646))
        # 选定后点击确定按钮位置
        btn_select_sure_location = Point(point_tuple=(480, 857))

        # one window
        # 方式选择按钮位置
        x_aix = 815
        y_aix = 530
        btn_method_location = Point(point_tuple=(x_aix, y_aix))
        # 方式选定按钮位置
        y_offset = 30
        btn_method_selected_location = Point(point_tuple=(x_aix, y_aix + y_offset))
        # 弹出选定条目窗口按钮位置
        y_offset = 95
        btn_pop_item_select_location = Point(point_tuple=(1010, y_aix + y_offset))

        # one window
        # 选择条目
        x_aix = 203
        y_aix = 500
        btn_item_select_location = Point(point_tuple=(x_aix, y_aix))
        # 选定条目，确认按钮
        y_offset = 304
        btn_item_select_sure_location = Point(point_tuple=(x_aix, y_aix + y_offset))

        # 输入金额输入框位置
        inbx_amount_location = Point(point_tuple=(781, 586))
        # 确认完成
        btn_finish_location = Point(point_tuple=(545, 876))

        combination_list = []
        # 点击新增按钮
        btn_new_apply = ButtonClick(btn_new_apply_location)
        combination_list.append(btn_new_apply)
        # 点击学生名称旁边的按钮，通过所选字段搜索
        btn_pop_search_student = ButtonClick(btn_pop_search_student_window_location)
        combination_list.append(btn_pop_search_student)
        # 向id输入框输入id
        inbx_id = InputBox(inbx_id_location, id)
        combination_list.append(inbx_id)
        # 点击查询按钮查询
        btn_search_student = ButtonClick(btn_search_student_location)
        combination_list.append(btn_search_student)
        # 选择对应学生
        btn_specified_student = ButtonClick(btn_specified_student_location)
        combination_list.append(btn_specified_student)
        # 选定后点击确定
        btn_select_sure = ButtonClick(btn_select_sure_location)
        combination_list.append(btn_select_sure)

        # 方式选择
        btn_method = ButtonClick(btn_method_location, duration=1)
        # 方式选定
        btn_method_selected = ButtonClick(btn_method_selected_location, duration=3600)
        # 组合前两项操作, 若不组合，则第一个会失去焦点
        combination_two_btn_method = Combination([btn_method, btn_method], duration=3)
        combination = Combination([combination_two_btn_method, btn_method_selected], duration=3)
        #operation_list.append(combination)

        operation_list.append(Combination(combination_list, duration=0.5))

        # 弹出选定条目
        btn_pop_item_select = ButtonClick(btn_pop_item_select_location)
        #operation_list.append(btn_pop_item_select)
        # 选择条目
        btn_item_select = ButtonClick(btn_item_select_location)
        #operation_list.append(btn_item_select)
        # 选定条目，确认按钮
        btn_item_select_sure = ButtonClick(btn_item_select_sure_location)
        #operation_list.append(btn_item_select_sure)

        # 输入金额
        inbx_amount = InputBox(inbx_amount_location, amount)
        operation_list.append(inbx_amount)
        # 确认完成
        btn_finish = ButtonClick(btn_finish_location)
        operation_list.append(btn_finish)
