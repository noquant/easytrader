# coding: utf-8
import os
import sys
import time
import unittest

sys.path.append(".")

TEST_CLIENTS = set(os.environ.get("EZ_TEST_CLIENTS", "").split(","))

IS_WIN_PLATFORM = sys.platform != "darwin"


@unittest.skipUnless("universal" in TEST_CLIENTS and IS_WIN_PLATFORM, "skip universal test")
class TestUniversalClientTrader(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import easytrader

        if "universal" not in TEST_CLIENTS:
            return

        # input your test account and password
        cls._ACCOUNT = os.environ.get("EZ_TEST_GTJA_ACCOUNT") or "your account"
        cls._PASSWORD = os.environ.get("EZ_TEST_GTJA_PASSWORD") or "your password"
        cls._COMM_PASSWORD = (
            os.environ.get("EZ_TEST_GTJA_COMM_PASSWORD") or "your comm password"
        )

        cls._user = easytrader.use("gtja_client")

        cls._user.prepare(
            user=cls._ACCOUNT, password=cls._PASSWORD, comm_password=cls._COMM_PASSWORD, exe_path="C:\\new_gtja_v6\\bin\\RichEZ.exe"
        )
        cls._user.enable_type_keys_for_editor()

    def test_balance_position(self):
        result1, result2 = self._user.balance_position
        print(result1, result2)

    def test_balance(self):
        result = self._user.balance
        print(result)

    def test_postion(self):
        result = self._user.position
        print(result)

    def test_today_entrusts(self):
        result = self._user.today_entrusts
        print(result)

    def test_today_trades(self):
        result = self._user.today_trades
        print(result)

    def test_cancel_entrusts(self):
        result = self._user.cancel_entrusts
        print(result)

    def test_cancel_entrust(self):
        result = self._user.cancel_entrust('249279')
        print(result)

    def test_invalid_market_buy(self):
        import easytrader

        with self.assertRaises(easytrader.exceptions.TradeError):
            result = self._user.market_buy("511990", 1e10)
            print(result)

    def test_invalid_market_sell(self):
        import easytrader

        with self.assertRaises(easytrader.exceptions.TradeError):
            result = self._user.market_sell("162411", 1e10)
            print(result)

    def test_invalid_buy(self):
        import easytrader

        with self.assertRaises(easytrader.exceptions.TradeError):
            result = self._user.buy("511990", 1, 1e10)
            print(result)

    def test_invalid_sell(self):
        import easytrader

        with self.assertRaises(easytrader.exceptions.TradeError):
            result = self._user.sell("162411", 200, 1e10)
            print(result)

    def test_auto_ipo(self):
        self._user.auto_ipo()


if __name__ == "__main__":
    unittest.main(verbosity=2)
