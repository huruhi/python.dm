from win32com.client import Dispatch


class DM:

    def __init__(self):

        try:

            self.dm = Dispatch('dm.dmsoft')

        except:

            import sys
            from time import sleep
            from os import system
            from os.path import abspath
            from os.path import join
            from tools import resource_path

            path = sys._MEIPASS if getattr(sys, 'frozen', False) else join(abspath('..'), 'dm')
            system('regsvr32 {} /s'.format(join(path, 'dm.dll')))
            self.dm = Dispatch('dm.dmsoft')

    @staticmethod
    def __list(value: str, convert=[str, int][0]) -> list:
        """
        将大漠返回的列表字符串转换为 list
        """

        if value:
            data = []
            for v in list(value.split(',')):
                data.append(convert(v))
            return data

        return []

    """
    窗口
    """

    def enum_window(self, parent: int, title: str, class_name: str, filter_mode: int) -> list:
        return self.__list(self.dm.EnumWindow(parent, title, class_name, filter_mode), int)

    def enum_window_by_process(self, process_name: str, title: str, class_name: str,
                               filter_mode: int = [1, 2, 4, 8, 16, 32][0]) -> list:
        return self.__list(self.dm.EnumWindowByProcess(process_name, title, class_name, filter_mode), int)

    def find_window(self, title: str, class_name: str = '') -> int:
        """
        查找符合类名或者标题名的顶层可见窗口

        :param title:       窗口标题，如果为空则模糊匹配匹配所有
        :param class_name:  窗口类名，如果为空则模糊匹配匹配所有
        """

        return self.dm.FindWindow(class_name, title)

    def find_window_by_process_id(self, process_id: int, class_name: str, title: str) -> int:
        """
        根据指定的 PID 查找可见窗口

        :param process_id:  PID
        :param class_name:  窗口类名
        :param title:       窗口标题
        """

        return self.dm.FindWindowByProcessId(process_id, class_name, title)

    def get_mouse_point_window(self):
        return self.dm.GetMousePointWindow()

    def get_window(self, hwnd: int, flag: int = [0, 1, 2, 3, 4, 5, 6, 7][0]) -> int:
        """
        获取给定窗口相关的窗口句柄

        :param hwnd:    窗口句柄
        :param flag:
        """

        return self.dm.GetWindow(hwnd, flag)

    def get_window_state(self, hwnd: int, flag: int = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9][0]) -> int:
        return self.dm.GetWindowState(hwnd, flag)

    def get_window_title(self, hwnd: int) -> str:
        return self.dm.GetWindowTitle(hwnd)

    def get_client_size(self, hwnd: int):
        w = None
        h = None
        return self.dm.GetClientSize(hwnd, w, h)

    def set_window_size(self, hwnd: int, width: int, height: int) -> int:
        return self.dm.SetWindowSize(hwnd, width, height)

    def set_window_state(self, hwnd: int, flag: int) -> int:
        return self.dm.SetWindowState(hwnd, flag)

    """
    后台
    """

    def bind_window(self, hwnd: int, display: str = ['normal', 'gdi', 'gdi2', 'dx', 'dx2'][0],
                    mouse: str = ['normal', 'windows', 'windows2', 'windows3', 'dx', 'dx2'][0],
                    keypad: str = ['normal', 'windows', 'dx'][0],
                    mode: int = [0, 1, 2, 3, 4, 5, 6, 7, 101, 103][0]) -> int:
        return self.dm.BindWindow(hwnd, display, mouse, keypad, mode)

    def bind_window_ex(self, hwnd: int, display: str, mouse: str, keypad: str, public: str, mode: int):
        return self.dm.BindWindowEx(hwnd, display, mouse, keypad, public, mode)

    def enable_mouse_sync(self, enable: int, time_out: int) -> int:
        return self.dm.EnableMouseSync(enable, time_out)

    def enable_real_mouse(self, enable: int, mousedelay: int, mousestep: int) -> int:
        return self.dm.EnableRealMouse(enable, mousedelay, mousestep)

    def un_bind_window(self) -> int:
        return self.dm.UnBindWindow()

    """
    基本设置
    """

    def get_id(self) -> int:
        return self.dm.GetID()

    def set_path(self, path: str) -> int:
        return self.dm.SetPath(path)

    """
    键鼠
    """

    def get_cursor_pos(self):
        x = None
        y = None
        return self.dm.GetCursorPos(x, y)

    def left_click(self) -> int:
        return self.dm.LeftClick()

    def left_up(self) -> int:
        return self.dm.LeftUp()

    def move_to(self, x: int, y: int) -> int:
        return self.dm.MoveTo(x, y)

    def move_to_ex(self, x: int, y: int, w: int, h: int):
        return self.dm.MoveToEx(x, y, w, h)

    """
    图色
    """

    def capture(self, x1: int, y1: int, x2: int, y2: int, file: str) -> int:
        return self.dm.Capture(x1, y1, x2, y2, file)

    def capture_png(self, x1: int, y1: int, x2: int, y2: int, file: str) -> int:
        return self.dm.CapturePng(x1, y1, x2, y2, file)

    def find_color(self, x1: int, y1: int, x2: int, y2: int, color: str, sim: float, dir_mode: int) -> int:
        int_x = int
        int_y = int
        return self.dm.FindColor(x1, y1, x2, y2, color, sim, dir_mode, int_x, int_y)

    """
    文字识别
    """

    def add_dict(self, index: int, dict_info: str) -> int:
        return self.dm.AddDict(index, dict_info)

    def fetch_word(self, x1: int, y1: int, x2: int, y2: int, color: str, word: str):
        return self.dm.FetchWord(x1, y1, x2, y2, color, word)

    def ocr(self, x1: int, y1: int, x2: int, y2: int, color_format: str, sim: float) -> str:
        return self.dm.Ocr(x1, y1, x2, y2, color_format, sim)

    def save_dict(self, index: int, file: str) -> int:
        return self.dm.SaveDict(index, file)

    def set_dict(self, index: int, file: str) -> int:
        return self.dm.SetDict(index, file)

    def use_dict(self, index: int) -> int:
        return self.dm.UseDict(index)

    """
    系统
    """

    def beep(self, f: int, duration: int) -> int:
        """
        蜂鸣器

        :param f:           频率
        :param duration:    时长
        """

        return self.dm.Beep(f, duration)

    def check_font_smooth(self) -> int:
        return self.dm.CheckFontSmooth()

    # @classmethod
    # def enum_process(cls, name: str = None) -> list:
    #     """
    #     根据指定条件，枚举系统中符合条件的进程
    #
    #     :param name:    进程名称，严格匹配
    #     """
    #
    #     pids = []
    #     for p in psutil.process_iter(['pid', 'name']):
    #         if p.info['name'] == name or not name:
    #             pids.append(p.info['pid'])
    #
    #     return pids

    def delay(self, mis: int) -> int:
        return self.dm.Delay(mis)

    def delays(self, mis_min: int, mis_max: int) -> int:
        return self.dm.Delays(mis_min, mis_max)

    def get_last_error(self) -> int:
        """
        获取插件命令的最后错误
        """

        return self.dm.GetLastError()

    def reg(self, reg_code: str, ver_info: str) -> int:
        """
        调用此函数来注册
        """

        return self.dm.Reg(reg_code, ver_info)

    def ver(self) -> str:
        """
        返回当前插件版本号
        """

        return self.dm.ver()

# dm = win32com.client.Dispatch('dm.dmsoft')  # 调用大漠插件
# print(dm.ver())  # 输出版本号
# print(dm.GetID())  # 输出当前大漠对象 ID
