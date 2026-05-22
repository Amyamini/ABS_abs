from typing import List, Dict
import datetime


class FundProduct:
    def __init__(self, name: str):
        self.name = name
        # 使用字典来存储每日现金流，日期作为键，金额作为值
        self.cash_flows: Dict[datetime.date, float] = {}

    def add_cash_flow(self, date: datetime.date, amount: float):
        """添加新的现金流"""
        if date in self.cash_flows:
            raise ValueError("Date already has cash flow")
        self.cash_flows[date] = amount

    def get_balance(self, end_date: datetime.date) -> float:
        """获取到指定日期的余额"""
        balance = 0.0
        for date, amount in self.cash_flows.items():
            if date <= end_date:
                balance += amount
        return balance


class FundPortfolio:
    def __init__(self):
        # 使用列表来存储基金产品
        self.products: List[FundProduct] = []

    def add_product(self, product: FundProduct):
        """添加新的基金产品"""
        if product in self.products:
            raise ValueError("Product already exists")
        self.products.append(product)

    def get_portfolio_balance(self, end_date: datetime.date) -> float:
        """获取到指定日期的组合余额"""
        balance = 0.0
        for product in self.products:
            balance += product.get_balance(end_date)
        return balance


```



```python
import unittest
from datetime import date, timedelta


class TestFundProduct(unittest.TestCase):
    def setUp(self):
        self.product = FundProduct('Test Product')

    def test_add_cash_flow(self):
        today = date.today()
        tomorrow = today + timedelta(days=1)
        self.product.add_cash_flow(tomorrow, 100.0)
        self.assertEqual(self.product.get_balance(tomorrow), 100.0)

    def test_get_balance(self):
        today = date.today()
        tomorrow = today + timedelta(days=1)
        self.product.add_cash_flow(today, -50.0)
        self.assertEqual(self.product.get_balance(tomorrow), 50.0)


class TestFundPortfolio(unittest.TestCase):
    def setUp(self):
        self.portfolio = FundPortfolio()
        self.product1 = FundProduct('Test Product 1')
        self.product2 = FundProduct('Test Product 2')

    def test_add_product(self):
        self.portfolio.add_product(self.product1)
        self.assertIn(self.product1, self.portfolio.products)

    def test_get_portfolio_balance(self):
        today = date.today()
        tomorrow = today + timedelta(days=1)
        self.product1.add_cash_flow(tomorrow, 100.0)
        self.product2.add_cash_flow(tomorrow, -50.0)
        self.portfolio.add_product(self.product1)
        self.portfolio.add_product(self.product2)
        self.assertEqual(self.portfolio.get_portfolio_balance(tomorrow), 50.0)


if __name__ == '__main__':
    unittest.main()