from operation_component import *
from PyQt5.QtWidgets import QWidget, QPushButton
from PyQt5.QtCore import Qt
import pyautogui


class MainWindow(QWidget):

    def __init__(self, height, width, location, title):
        super().__init__()
        self._operation_list = []
        self.add_all_operation()
        self._current_operation_index = 0
        self._store_position = "asdfd"
        self._start_btn = None
        self.initUI(height, width, location, title)
        self._is_started = False

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

        # 窗体总在最前端
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.show()

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
        if 0 < self._current_operation_index < len(self._operation_list):
            index = self._current_operation_index
            self._current_operation_index += 1
            self.one_operation(self._operation_list[index])
            index = self._current_operation_index
            if index == 8:
                self._current_operation_index += 1
                self.one_operation(self._operation_list[index])
        elif self._current_operation_index == 0:
            print("请先开始")
        else:
            print("执行完毕")

    def last_operation(self):
        if self._current_operation_index == 9:
            self._current_operation_index -= 2
        elif self._current_operation_index > 0:
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
            with open("current_location.txt", "w", encoding='utf-8') as f:
                f.write(self._store_position)

    def add_all_operation(self):
        self._store_position = u"孔祥富"
        amount = "600"
        id = ""
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
        y_aix = 530
        btn_method_location = Point(point_tuple=(817, y_aix))
        # 方式选定按钮位置
        offset = 37
        btn_method_selected_location = Point(point_tuple=(817, y_aix + offset))
        # 弹出选定条目窗口按钮位置
        offset = 95
        btn_pop_item_select_location = Point(point_tuple=(1010, y_aix + offset))

        # one window
        # 选择条目
        x_aix = 203
        y_aix = 500
        btn_item_select_location = Point(point_tuple=(x_aix, y_aix))
        # 选定条目，确认按钮
        offset = 304
        btn_item_select_sure_location = Point(point_tuple=(x_aix, y_aix + 304))

        # 输入金额输入框位置
        inbx_amount_location = Point(point_tuple=(781, 586))
        # 确认完成
        btn_finish_location = Point(point_tuple=(545, 876))

        # 点击新增按钮
        btn_new_apply = ButtonClick(btn_new_apply_location)
        operation_list.append(btn_new_apply)
        # 点击学生名称旁边的按钮，通过所选字段搜索
        btn_pop_search_student = ButtonClick(btn_pop_search_student_window_location)
        operation_list.append(btn_pop_search_student)
        # 向id输入框输入id
        inbx_id = InputBox(inbx_id_location, id)
        operation_list.append(inbx_id)
        # 点击查询按钮查询
        btn_search_student = ButtonClick(btn_search_student_location)
        operation_list.append(btn_search_student)
        # 选择对应学生
        btn_specified_student = ButtonClick(btn_specified_student_location)
        operation_list.append(btn_specified_student)
        # 选定后点击确定
        btn_select_sure = ButtonClick(btn_select_sure_location)
        operation_list.append(btn_select_sure)

        # 方式选择
        btn_method = ButtonClick(btn_method_location, duration=0.2)
        operation_list.append(btn_method)
        operation_list.append(btn_method)
        # 方式选定
        btn_method_selected = ButtonClick(btn_method_selected_location)
        operation_list.append(btn_method_selected)

        # 弹出选定条目
        btn_pop_item_select = ButtonClick(btn_pop_item_select_location)
        operation_list.append(btn_pop_item_select)
        # 选择条目
        btn_item_select = ButtonClick(btn_item_select_location)
        operation_list.append(btn_item_select)
        # 选定条目，确认按钮
        btn_item_select_sure = ButtonClick(btn_item_select_sure_location)
        operation_list.append(btn_item_select_sure)
        # 输入金额
        inbx_amount = InputBox(inbx_amount_location, amount)
        operation_list.append(inbx_amount)
        # 确认完成
        # btn_finish = ButtonClick(btn_finish_location)
        # operation_list.append(btn_finish)
