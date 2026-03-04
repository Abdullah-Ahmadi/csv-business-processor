import pandas as pd

from .base_processor import BaseProcessor
from utils.formatters import format_percentage, format_currency


class FinanceProcessor(BaseProcessor):

    def load_data(self, filepath):

        self.logger.info(f"Loding financial data from {filepath}")

        try:
            self.data = pd.read_csv(filepath)
            self._standardize_columns_names()

            if "date" in self.data.columns:
                self.data["date"] = pd.to_datetime(self.data["date"], errors="coerce")

            if "amount" in self.data.columns:
                self.data["amount"] = pd.to_numeric(
                    self.data["amount"], errors="coerce"
                )

                self.data["amount"] = self.data["amount"].abs()

            if "category" in self.data.columns:
                self.data["category"] = (
                    self.data["category"].str.lower().str.strip()
                )  # faulty with line 180, no category

            self.logger.info(f"Loaded {len(self.data)} financial transactions")
            return self

        except Exception as e:
            self.logger.info(f"Error loading financial data: {e}")
            raise

    def analyze(self):
        self.logger.info("Analyzing financial data...")
        self.validate()

        self.insights = {}
        self.alerts = []

        if "amount" not in self.data.columns:
            raise ValueError("Missing required column: 'amount'")

        # Basic Metrix
        self.insights["total_spent"] = float(self.data["amount"].sum())
        self.insights["transaction_count"] = int(len(self.data))
        self.insights["average_transaction"] = float(
            self.insights["total_spent"] / self.insights["transaction_count"]
            if self.insights["transaction_count"] > 0
            else 0
        )

        # Category analysis
        if "category" in self.data.columns:
            category_spending = (
                self.data.groupby("category")["amount"]
                .sum()
                .sort_values(ascending=False)
            )
            self.insights["spending_by_category"] = category_spending.to_dict()
            self.insights["top_category"] = (
                category_spending.index[0] if not category_spending.empty else "N/A"
            )
            self.insights["top_category_amount"] = float(
                category_spending.iloc[0] if not category_spending.empty else 0
            )

            # Calculate percentage breakdown
            if not self.insights["total_spent"] > 0:
                for category, amount in category_spending.items():
                    percentage = (amount / self.insights["total_spent"]) * 100
                    if percentage > 30:
                        self.alerts.append(
                            f"High spending in '{category}': {format_percentage(percentage)} of total ({format_currency(amount)})"
                        )

        # Time based analysis
        if "date" in self.data.columns and not self.data["date"].isnull().all():
            self.data["month"] = self.data["date"].dt.strftime("%Y-%m")
            monthly_spending = self.data.groupby("month")["amount"].sum()

            if not monthly_spending.empty:
                self.insights["monthly_spending"] = monthly_spending.to_dict()
                self.insights["highest_spending_month"] = str(monthly_spending.idxmax())
                self.insights["lowest_spending_month"] = str(monthly_spending.idxmin())
                self.insights["average_monthly_spending"] = float(
                    monthly_spending.mean()
                )

                # Monthly spending growth
                if len(monthly_spending) >= 2:
                    months = sorted(monthly_spending.index)
                    current = monthly_spending[months[-1]]
                    previous = monthly_spending[months[-2]]

                    if previous > 0:
                        growth = ((current - previous) / previous) * 100
                        self.insights["monthly_sepending_growth"] = growth

                        if growth > 20:
                            self.alerts.append(
                                f"Spending increased by {format_percentage(growth)} compared to previous month"
                            )

                        elif growth < -30:
                            self.alerts.append(
                                f"Spending decreased by {format_percentage(abs(growth))} - check for missing transactions"
                            )

        self.data["day_of_week"] = self.data["date"].dt.day_name()
        daily_spending = self.data.groupby("day_of_week")["amount"].sum()

        if not daily_spending.empty:
            self.insights["spending_by_day"] = daily_spending.to_dict()
            self.insights["highest_spending_day"] = daily_spending.idxmax()

        # Payment method analysis
        if "payment_method" in self.data.columns:
            payments_stats = (
                self.data.groupby("payment_method")
                .agg({"amount": ["sum", "count", "mean"]})
                .round(2)
            )

            payments_stats.columns = ["total", "count", "average"]
            self.insights["payment_method_stats"] = payments_stats.to_dict("index")

            if self.insights["total_spent"] > 0:
                for method, stats in payments_stats.iterrows():
                    percentage = (stats["total"] / self.insights["total_spent"]) * 100
                    if percentage > 60:
                        self.alerts.append(
                            f"Payment method '{method}' accounts for {format_percentage(percentage)} of total spending"
                        )

        # Merchant analytics
        if "merchant" in self.data.columns:
            merchant_spending = (
                self.data.groupby("merchant")["amount"]
                .sum()
                .sort_values(ascending=False)
            )

            top_merchants = merchant_spending.head(5)
            self.insights["top_merchant"] = top_merchants.to_dict()

            if not merchant_spending.empty and self.insights["total_spent"] > 0:
                top_merchant_percentage = (
                    merchant_spending.iloc[0] / self.insights["total_spent"]
                )

                if top_merchant_percentage > 40:
                    self.alerts.append(
                        f"High concentration: {merchant_spending.index[0]} accounts for "
                        f"{format_percentage(top_merchant_percentage)} of spending"
                    )
        if self.insights["total_spent"] > 0:
            large_transactions = self.data[
                (self.data["amount"] > 500)
                | (self.data["amount"] > self.insights["total_spent"] * 0.1)
            ]

            if not large_transactions.empty:
                self.insights["large_transactions"] = large_transactions[
                    ["date", "description", "amount", "category"]
                ].to_dict("records")

                for _, trans in large_transactions.head(3).iterrows():
                    desc = trans.get("description", "Unknown")[:30]

                    self.alerts.append(
                        f"Large transaction: ${trans['amount']:,.2f} for '{desc}'"
                    )

        # Savings opportunities
        savings_suggestions = []

        discretionary_categories = ["entertainment", "dining", "shopping", "hobbies"]

        if "category" in self.data.columns:
            for category in discretionary_categories:
                if category in self.insights.get("spending_by_category", {}):
                    amount = self.insights["spending_by_category"][category]
                    if amount > 200:
                        potential_savings = amount * 0.15
                        savings_suggestions.append(
                            {
                                "category": category,
                                "current_spending": amount,
                                "suggested_reduction": potential_savings,
                                "reason": "Discretionary spending could be reduced",
                            }
                        )
        if savings_suggestions:
            self.insights["savings_opportunities"] = savings_suggestions
            total_potential_savings = sum(
                item["suggested_reduction"] for item in savings_suggestions
            )

            self.alerts.append(
                f"Potential monthly savings: {format_currency(total_potential_savings)} by reducing discretionary spending"
            )

        self.logger.info(
            f"Financial analysis complete: {len(self.insights)} insights, {len(self.alerts)} alerts"
        )
        return self

    def _standardize_column_names(self):
        """Standardize financial specific column names."""
        if self.data is not None:
            column_mapping = {
                "transaction_date": "date",
                "purchase_date": "date",
                "description": "description",
                "details": "description",
                "transaction_details": "description",
                "transaction_amount": "amount",
                "price": "amount",
                "cost": "amount",
                "expense_amount": "amount",
                "category": "category",
                "expense_category": "category",
                "payment_method": "payment_method",
                "payment_type": "payment_method",
                "merchant": "merchant",
                "vendor": "merchant",
                "store": "merchant",
            }

            self.data.columns = [
                str(col).strip().lower().replace(" ", "_") for col in self.data.columns
            ]
            self.data.columns = [
                column_mapping.get(col, col) for col in self.data.columns
            ]

    def validate(self):
        return super().validate()
