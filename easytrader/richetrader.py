# -*- coding: utf-8 -*-
import abc
import functools
import logging
import os
import re
import sys
import time
from typing import Type, Union
import xlrd

import hashlib, binascii

import easyutils
from pywinauto import findwindows, timings
from pywinauto import mouse, keyboard
import pandas as pd

from easytrader import grid_strategies, pop_dialog_handler, refresh_strategies
from easytrader.config import client
from easytrader.grid_strategies import IGridStrategy
from easytrader.log import logger
from easytrader.refresh_strategies import IRefreshStrategy
from easytrader.utils.misc import file2dict
from easytrader.utils.perf import perf_clock

if not sys.platform.startswith("darwin"):
    import pywinauto
    import pywinauto.clipboard

class IClientTrader(abc.ABC):
    @property
    @abc.abstractmethod
    def app(self):
        """Return current app instance"""
        pass

    @property
    @abc.abstractmethod
    def main(self):
        """Return current main window instance"""
        pass

    @property
    @abc.abstractmethod
    def config(self):
        """Return current config instance"""
        pass

    @abc.abstractmethod
    def wait(self, seconds: float):
        """Wait for operation return"""
        pass

    @abc.abstractmethod
    def is_exist_pop_dialog(self):
        pass


class RicheTrader(IClientTrader):
    _editor_need_type_keys = False
    # The strategy to use for getting grid data
    grid_strategy: Union[IGridStrategy, Type[IGridStrategy]] = grid_strategies.Xls97
    _grid_strategy_instance: IGridStrategy = None
    refresh_strategy: IRefreshStrategy = refresh_strategies.Panelbar(85)

    def enable_type_keys_for_editor(self):
        """
        有些客户端无法通过 set_edit_text 方法输入内容，可以通过使用 type_keys 方法绕过
        """
        self._editor_need_type_keys = True

    @property
    def grid_strategy_instance(self):
        if self._grid_strategy_instance is None:
            self._grid_strategy_instance = (
                self.grid_strategy
                if isinstance(self.grid_strategy, IGridStrategy)
                else self.grid_strategy()
            )
            self._grid_strategy_instance.set_trader(self)
        return self._grid_strategy_instance

    def __init__(self):
        self._config = client.create(self.broker_type)
        self._app = None
        self._main = None
        self._toolbar = None

    @property
    def app(self):
        return self._app

    @property
    def main(self):
        return self._main

    @property
    def config(self):
        return self._config

    def connect(self, exe_path=None, **kwargs):
        """
        直接连接登陆后的客户端
        :param exe_path: 客户端路径类似 r'C:\\htzqzyb2\\xiadan.exe', 默认 r'C:\\htzqzyb2\\xiadan.exe'
        :return:
        """
        connect_path = exe_path or self._config.DEFAULT_EXE_PATH
        if connect_path is None:
            raise ValueError(
                "参数 exe_path 未设置，请设置客户端对应的 exe 地址,类似 C:\\客户端安装目录\\xiadan.exe"
            )

        self._app = pywinauto.Application().connect(path=connect_path, timeout=10)
        self._close_prompt_windows()
        self._main = self._app.top_window()
        self._init_toolbar()

    @property
    def broker_type(self):
        return "ths"

    def _get_data_from_panel(self, panel, title, refresh=False):
        result = []
        rect = panel.rectangle()
        if refresh:
            x, y = rect.right - 85, (rect.top + rect.bottom) // 2
            mouse.move(coords=(x, y))
            mouse.click(coords=(x, y))
            time.sleep(0.5)
        x, y = rect.right - 20, (rect.top + rect.bottom) // 2
        mouse.move(coords=(x, y))
        mouse.click(coords=(x, y))
        time.sleep(1.5)
        self.app.top_window().type_keys("^C", set_foreground=False)
        self.app.top_window().type_keys("%{s}%{y}", set_foreground=False)
        time.sleep(1)
        #
        try:
            file_name = pywinauto.clipboard.GetData()
            #
            data = xlrd.open_workbook('C:\\Users\\Administrator\\Desktop\\%s' % file_name, encoding_override='GBK')
            table = data.sheets()[0]
            #
            nrows = table.nrows
            for i in range(nrows):
                # print(table.row_values(i))
                result.append(table.row_values(i))
        # pylint: disable=broad-except
        except Exception as e:
            print(e)
        #
        return result

    def _to_list_dict(self, data):
        df = pd.DataFrame(
            data[1:],
            columns=data[0]
        )
        return df.to_dict("records")

    @property
    def balance(self):
        self._switch_left_menus_by_shortcut("{F4}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        data = self._get_grid_data(panel.TspSkinPanel6.control_id())
        return self._to_list_dict(data[5:7])[0]

    @property
    def balance_position(self):
        self._switch_left_menus_by_shortcut("{F4}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        self.refresh(panel.TspSkinPanel6)
        data = self._get_grid_data(panel.TspSkinPanel6.control_id())
        return self._to_list_dict(data[5:7])[0], self._to_list_dict(data[8:-1])

    def _init_toolbar(self):
        self._toolbar = self._main.child_window(class_name="ToolbarWindow32")

    @property
    def position(self):
        self._switch_left_menus_by_shortcut("{F4}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        data = self._get_grid_data(panel.TspSkinPanel6.control_id())
        return self._to_list_dict(data[8:-1])

    @property
    def today_entrusts(self):
        self._switch_left_menus_by_shortcut("{F7}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        data = self._get_grid_data(panel.TspSkinPanel5.control_id())
        return self._to_list_dict(data[5:])

    @property
    def today_trades(self):
        self._switch_left_menus_by_shortcut("{F6}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        data = self._get_grid_data(panel.TspSkinPanel5.control_id())
        return self._to_list_dict(data[5:])

    @property
    def cancel_entrusts(self):
        self._switch_left_menus_by_shortcut("{F8}")

        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        # panel.print_ctrl_ids()
        data = self._get_grid_data(panel.TspSkinPanel5.control_id())
        return self._to_list_dict(data[5:])

    @perf_clock
    def cancel_entrust(self, entrust_no):
        for i, entrust in enumerate(self.cancel_entrusts):
            if entrust['委托号'] == entrust_no:
                self._cancel_entrust_by_double_click(i)
                # 如果出现了确认窗口
                self.close_pop_dialog()
        return {"message": "委托单状态错误不能撤单, 该委托单可能已经成交或者已撤"}

    def cancel_all_entrusts(self):
        self._switch_left_menus_by_shortcut("{F8}A")
        self.wait(0.2)
        # 等待出现 确认兑换框
        if self.is_exist_pop_dialog():
            # 点击是 按钮
            w = self._app.top_window()
            if w is not None:
                btn = w["确定"]
                if btn is not None:
                    btn.click()
                    self.wait(0.2)
        # 如果出现了确认窗口
        self.close_pop_dialog()

    @perf_clock
    def buy(self, security, price, amount, **kwargs):
        self._switch_left_menus_by_shortcut("{F2}")

        return self.trade('限价买入(&B)', security, price, amount)

    @perf_clock
    def sell(self, security, price, amount, **kwargs):
        self._switch_left_menus_by_shortcut("{F3}")

        return self.trade('限价卖出(&S)', security, price, amount)

    @perf_clock
    def market_buy(self, security, amount, ttype=None, limit_price=None, **kwargs):
        self._switch_left_menus_by_shortcut("{F2}")

        return self.market_trade('限价买入(&B)', security, amount)

    @perf_clock
    def market_sell(self, security, amount, ttype=None, limit_price=None, **kwargs):
        self._switch_left_menus_by_shortcut("{F3}")

        return self.market_trade('限价卖出(&S)', security, amount)

    def market_trade(self, title, security, amount, ttype=None, limit_price=None, **kwargs):
        self._set_market_trade_params(title, security, amount)

        self._submit_trade(title)

        return self._handle_pop_dialogs(
            handler_class=pop_dialog_handler.EnterDialogHandler
        )

    def auto_ipo(self):
        self._switch_left_menus(["新股申购", "今日申购"])
        #
        btn = self._main.child_window(
            title='批量申购', class_name="TspSkinButton"
        )
        rect = btn.rectangle()
        x, y = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
        mouse.move(coords=(x, y))
        mouse.click(coords=(x, y))
        self.wait(1)
        self._app.top_window().type_keys('%{U}')
        self.wait(1)
        btn = self._app.top_window().child_window(
            title='确认申购', class_name="TspSkinButton"
        )
        rect = btn.rectangle()
        x, y = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
        mouse.move(coords=(x, y))
        mouse.click(coords=(x, y))
       
        return self._handle_pop_dialogs(
            handler_class=pop_dialog_handler.EnterDialogHandler
        )

    def _click_grid_by_row(self, row):
        x = self._config.COMMON_GRID_LEFT_MARGIN
        y = (
            self._config.COMMON_GRID_FIRST_ROW_HEIGHT
            + self._config.COMMON_GRID_ROW_HEIGHT * row
        )
        self._app.top_window().child_window(
            control_id=self._config.COMMON_GRID_CONTROL_ID,
            class_name="CVirtualGridCtrl",
        ).click(coords=(x, y))

    @perf_clock
    def is_exist_pop_dialog(self):
        self.wait(0.5)  # wait dialog display
        try:
            return (
                self._main.wrapper_object() != self._app.top_window().wrapper_object()
            )
        except (
            findwindows.ElementNotFoundError,
            timings.TimeoutError,
            RuntimeError,
        ) as ex:
            logger.exception("check pop dialog timeout")
            return False

    @perf_clock
    def close_pop_dialog(self):
        try:
            if self._main.wrapper_object() != self._app.top_window().wrapper_object():
                w = self._app.top_window()
                if w is not None:
                    w.close()
                    self.wait(0.2)
        except (
                findwindows.ElementNotFoundError,
                timings.TimeoutError,
                RuntimeError,
        ) as ex:
            pass

    def _run_exe_path(self, exe_path):
        return os.path.join(os.path.dirname(exe_path), "xiadan.exe")

    def wait(self, seconds):
        time.sleep(seconds*0.5)

    def exit(self):
        self._app.kill()

    def _close_prompt_windows(self):
        self.wait(1)
        for window in self._app.windows(class_name="#32770", visible_only=True):
            title = window.window_text()
            if title != self._config.TITLE:
                logging.info("close " + title)
                window.close()
                self.wait(0.2)
        self.wait(1)

    def close_pormpt_window_no_wait(self):
        for window in self._app.windows(class_name="#32770"):
            if window.window_text() != self._config.TITLE:
                window.close()

    def trade(self, title, security, price, amount):
        for _ in range(10):
            try:
                btn = self._main.child_window(title=title, class_name="TspSkinButton")
                btn.print_ctrl_ids()
            except Exception as e:
                print(e)
                self.wait(0.1)
        self._set_trade_params(security, price, amount)

        self._submit_trade(title)

        return self._handle_pop_dialogs(
            handler_class=pop_dialog_handler.EnterDialogHandler
        )

    def _click(self, control_id):
        self._app.top_window().child_window(
            control_id=control_id, class_name="Button"
        ).click()

    @perf_clock
    def _submit_trade(self, title):
        time.sleep(0.01)
        btn = self._main.child_window(title=title, class_name="TspSkinButton")
        btn.click()

    @perf_clock
    def __get_top_window_pop_dialog(self):
        return self._app.top_window().window(
            control_id=self._config.POP_DIALOD_TITLE_CONTROL_ID
        )

    @perf_clock
    def _get_pop_dialog_title(self):
        return (
            self._app.top_window().window_text()
        )

    def _set_trade_params(self, security, price, amount):
        code = security[-6:]
        editor = self._main.child_window(class_name="TspCustomEdit")
        editor.select()
        editor.type_keys(code)
        #
        panel = self._app.window(handle=editor.parent().parent().handle)
        # panel.print_ctrl_ids()
        # wait security input finish
        for _ in range(30):
            self.wait(0.1)
            editor = panel.Edit4
            texts = editor.texts()
            if len(texts) > 0 and len(texts[0]) > 0:
                break
        editor.select()
        str_bs = '{BACKSPACE}' * len(texts[0])
        editor.type_keys(str_bs)
        editor.type_keys(str(int(price)))
        #
        editor = panel.Edit2
        editor.select()
        editor.type_keys(str(int(amount)))


    def _set_market_trade_params(self, title, security, amount, limit_price=None):
        code = security[-6:]
        editor = self._main.child_window(class_name="TspCustomEdit")
        editor.select()
        editor.type_keys(code)
        #
        panel = self._app.window(handle=editor.parent().parent().handle)
        # panel.print_ctrl_ids()
        # wait security input finish
        for _ in range(30):
            self.wait(0.1)
            editor = panel.Edit4
            texts = editor.texts()
            if len(texts) > 0 and len(texts[0]) > 0:
                break
        editor.select()
        if '买' in title:
            editor.type_keys('{UP}{UP}{UP}')
        else:
            editor.type_keys('{DOWN}{DOWN}{DOWN}')
        #
        editor = panel.Edit2
        editor.select()
        editor.type_keys(str(int(amount)))

    def _get_grid_data(self, control_id):
        return self.grid_strategy_instance.get(control_id)

    def _type_keys(self, control_id, text):
        self._main.child_window(control_id=control_id, class_name="Edit").set_edit_text(
            text
        )

    def _type_edit_control_keys(self, control_id, text):
        if not self._editor_need_type_keys:
            self._main.child_window(
                control_id=control_id, class_name="Edit"
            ).set_edit_text(text)
        else:
            editor = self._main.child_window(control_id=control_id, class_name="Edit")
            editor.select()
            editor.type_keys(text)

    def type_edit_control_keys(self, editor, text):
        if not self._editor_need_type_keys:
            editor.set_edit_text(text)
        else:
            editor.select()
            editor.type_keys(text)

    def _collapse_left_menus(self):
        items = self._get_left_menus_handle().roots()
        for item in items:
            item.collapse()

    @perf_clock
    def _switch_left_menus(self, path, sleep=0.2):
        self.close_pop_dialog()
        self._app.top_window().type_keys('{F8}')
        rect = self._get_left_menus_handle().rectangle()
        x, y = (rect.left + rect.right) // 2, rect.top + 100
        mouse.move(coords=(x, y))
        mouse.click(coords=(x, y))
        self.wait(sleep)
        z = y + 30
        mouse.move(coords=(x, z))
        mouse.click(coords=(x, z))
        self.wait(sleep)
        try:
            btn = self._main.child_window(
                title='批量申购', class_name="TspSkinButton"
            )
            btn.print_ctrl_ids()
        except Exception as e:
            print('异常', e)
            mouse.move(coords=(x, y))
            mouse.click(coords=(x, y))
            self.wait(sleep)
            mouse.move(coords=(x, z))
            mouse.click(coords=(x, z))
            self.wait(sleep)

    def _switch_left_menus_by_shortcut(self, shortcut, sleep=0.5):
        self.close_pop_dialog()
        self._app.top_window().type_keys(shortcut)
        self.wait(sleep)

    @functools.lru_cache()
    def _get_left_menus_handle(self):
        count = 2
        while True:
            try:
                handle = self._main.child_window(
                    title='', class_name="TChildMenuBox"
                )
                if count <= 0:
                    return handle
                # sometime can't find handle ready, must retry
                handle.wait("ready", 2)
                return handle
            # pylint: disable=broad-except
            except Exception as ex:
                logger.exception("error occurred when trying to get left menus")
            count = count - 1

    def _cancel_entrust_by_double_click(self, row):
        x = self._config.CANCEL_ENTRUST_GRID_LEFT_MARGIN
        y = (
            self._config.CANCEL_ENTRUST_GRID_FIRST_ROW_HEIGHT
            + self._config.CANCEL_ENTRUST_GRID_ROW_HEIGHT * row
        )
        panel = self._main.child_window(title='bgpanel', class_name='TspSkinPanel')
        panel.child_window(title='', class_name="TAdvStringGrid").double_click(coords=(x, y))
        self.wait(0.2)
        # 等待出现 确认兑换框
        if self.is_exist_pop_dialog():
            # 点击是 按钮
            w = self._app.top_window()
            if w is not None:
                btn = w["确定"]
                if btn is not None:
                    btn.click()
                    rc = btn.rectangle()
                    x, y = (rc.left+rc.right)//2, (rc.top + rc.bottom)//2
                    mouse.move(coords=(x, y))
                    mouse.click(coords=(x, y))
                    self.wait(0.2)

    def refresh(self, panel):
        self.refresh_strategy.set_trader(self)
        self.refresh_strategy.set_panel(panel)
        self.refresh_strategy.refresh()

    @perf_clock
    def _handle_pop_dialogs(self, handler_class=pop_dialog_handler.PopDialogHandler):
        handler = handler_class(self._app)

        while self.is_exist_pop_dialog():
            try:
                title = self._get_pop_dialog_title()
            except pywinauto.findwindows.ElementNotFoundError:
                return {"message": "success"}

            result = handler.handle(title)
            if result:
                return result
        return {"message": "success"}


class BaseLoginClientTrader(RicheTrader):
    @abc.abstractmethod
    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """Login Client Trader"""
        pass

    def prepare(
        self,
        config_path=None,
        user=None,
        password=None,
        exe_path=None,
        comm_password=None,
        **kwargs
    ):
        """
        登陆客户端
        :param config_path: 登陆配置文件，跟参数登陆方式二选一
        :param user: 账号
        :param password: 明文密码
        :param exe_path: 客户端路径类似 r'C:\\htzqzyb2\\xiadan.exe', 默认 r'C:\\htzqzyb2\\xiadan.exe'
        :param comm_password: 通讯密码
        :return:
        """
        if config_path is not None:
            account = file2dict(config_path)
            user = account["user"]
            password = account["password"]
            comm_password = account.get("comm_password")
            exe_path = account.get("exe_path")
        self.login(
            user,
            password,
            exe_path or self._config.DEFAULT_EXE_PATH,
            comm_password,
            **kwargs
        )
        self._init_toolbar()
