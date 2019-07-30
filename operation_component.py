import pyautogui
from abc import ABCMeta, abstractmethod
import time

class Point:
    def __init__(self, x=0, y=0, point_tuple=None):
        if point_tuple is not None:
            self._x = point_tuple[0]
            self._y = point_tuple[1]
        else:
            self._x = x
            self._y = y

    @property
    def x(self):
        return self._x

    @property
    def y(self):
        return self._y

    def get_point_tuple(self):
        return tuple([self._x, self._y])

    @x.setter
    def x(self, x):
        self._x = x

    @y.setter
    def y(self, y):
        self._y = y


class Component(metaclass=ABCMeta):
    @abstractmethod
    def operate(self):
        pass


class ButtonClick(Component):
    def __init__(self, location, target_img=None, is_click=True, duration=0):
        self._button_location = location
        self._target_img = target_img
        self._is_click = is_click
        self._duration = duration

    def operate(self):
        self.click()

    def click(self):
        # 这里不考虑坐标及目标截图同时输入的情况
        if self._target_img is not None:
            # 这里暂时只考虑只对搜索到的第一处进行处理，后期考虑多处情况
            button_locations = pyautogui.locateOnScreen(self._target_img)
            if button_locations is not None:
                button_locations_list = list(button_locations)
                # 拿到第一个位置的四整数元组
                button_location = button_locations_list[0]
                # 获得按钮的中心
                button_center = pyautogui.center(button_location)
                self._button_location.x(button_center[0])
                self._button_location.y(button_center[1])
            else:
                raise KeyboardInterrupt("目标截图不存在")

        elif self._button_location is not None:
            x = self._button_location.x
            y = self._button_location.y
            pyautogui.moveTo(x, y)
            max_time = 3600
            click = 0
            if self._is_click:
                if self._duration == click:
                    pyautogui.click(x, y)
                elif self._duration == max_time:
                    pyautogui.mouseDown(x, y)
                else:
                    pyautogui.mouseDown(x, y)
                    time.sleep(self._duration)
                    pyautogui.mouseUp(x, y)
            else:
                pyautogui.doubleClick(x, y)
        else:
            raise KeyboardInterrupt("输入源不正确")

    @property
    def is_click(self):
        return self._is_click

    @is_click.setter
    def is_click(self, click):
        self._is_click = click



class InputBox(Component):

    def __init__(self, location, input_content):
        self._location = location
        self._input_content = input_content

    def operate(self):
        self.type_into(self._input_content)

    def type_into(self, input_content):
        x = self._location.x
        y = self._location.y
        # 移动到固定位置
        pyautogui.moveTo(x, y)
        # 使输入框获取焦点
        pyautogui.click(x, y)
        # 向输入框中输入字符串
        pyautogui.typewrite(input_content)

    @property
    def input_content(self, input_content):
        self._input_content = input_content

    @input_content.setter
    def input_content(self, input_content):
        self._input_content = input_content
