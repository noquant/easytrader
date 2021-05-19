# -*- coding: utf-8 -*-

import pywinauto
import pywinauto.clipboard

from pywinauto import mouse, keyboard

from easytrader import grid_strategies
from easytrader.utils.win_gui import SetForegroundWindow
from . import richetrader


class GTJARicheTrader(richetrader.BaseLoginClientTrader):

    @property
    def broker_type(self):
        return "universal"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
        :param user: 用户名
        :param password: 密码
        :param exe_path: 客户端路径, 类似
        :param comm_password:
        :param kwargs:
        :return:
        """
        self._editor_need_type_keys = False

        try:
            self._app = pywinauto.Application().connect(
                path=exe_path, timeout=3
            )
            self._close_prompt_windows()
            lock_win = self._app.window(title=" 富易交易")
            if lock_win.exists() and lock_win.is_visible():
                lock_win.print_ctrl_ids()
                SetForegroundWindow(lock_win.wrapper_object())
                rect = lock_win.child_window(title="解锁", class_name="TspSkinButton").rectangle()
                x, y = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
                mouse.move(coords=(x, y))
                mouse.click(coords=(x, y))
                self.wait(1)
                passedit = self._app.window(title="-%s" % user).child_window(title="", class_name="TFyPassEdit")
                passedit.select()
                for k in password:
                    keyboard.send_keys('{VK_NUMPAD%s}' % k)
                keyboard.send_keys('{ENTER}')
            self.wait(1)
            self._main = self._app.window(title="富易 - %s" % user, class_name='TMainForm')
            self.wait(1)
        # pylint: disable=broad-except
        except Exception as e:
            self._app = pywinauto.Application().start(exe_path)
            self.wait(1)
