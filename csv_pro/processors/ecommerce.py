import pandas as pd
from .base_processor import BaseProcessor


class EcommerceProcessor(BaseProcessor):

    def load_data(self, filepath):
        self.logger.info(f"Loading e-commerce data from {filepath}")

        try:
            self.data = pd.read_csv(filepath)

            self._standardize_columns_names()

            if "order_date" in self.data.columns:
                self.data.order_date = pd.to_datetime(
                    self.data.order_date, errors="coerce"
                )

            numeric_columns = ["qauntity", "unit_price", "total_amount"]
            for col in numeric_columns:
                if col in self.data.columns:
                    self.data[col] = pd.to_numeric(self.data[col], errors="coerce")

            self.logger.info(
                f"Loaded {len(self.data)} rows with {len(self.data.columns)} columns"
            )
            return self

        except Exception as e:
            self.logger.info(f"Error loading e-commerce data: {e}")
            raise

    def analyze(self):

        self.logger.info(f"Analyzing e-commerce data...")

        self.validate()

        self.insights = {}
        self.alerts = []

        # Revenue analysis
        if "total_amount" in self.data.columns:
            self.insights["total_revenue"] = float(self.data["total_amount"].sum())
            self.insights["total_orders"] = int(self.data["order_id"].nunique())
            self.insights["average_order_value"] = float(
                self.insights["total_revenue"] / self.insights["total_orders"]
                if self.insights["total_orders"] > 0
                else 0
            )

        if "quantity" in self.data.columns:
            self.insights["total_units_sold"] = int(self.data["quantity"].sum())

        # Top and bottom products identification
        if "total_amount" in self.data.columns and "product_name" in self.data.columns:
            product_stats = (
                self.data.groupby("product_name")
                .agg({"total_amount": "sum", "order_id": "count", "quantity": "sum"})
                .round(2)
            )

            product_stats.columns = ["revenue", "units_sold", "order_count"]

            top_products = product_stats.sort_values("revenue", ascending=False).head(5)
            self.insights["top_products"] = top_products.to_dict("index")

            if len(product_stats) > 5:
                bottom_products = product_stats.sort_values("revenue").head(3)
                if not bottom_products.empty and bottom_products["revenue"].sum() < 100:
                    self.alerts.append(
                        f"Low revenue product detected: {', '.join(bottom_products.index[:2])}"
                    )

        # Time-based analysis
        if (
            "order_date" in self.data.columns
            and not self.data["order_date"].isnull().all()
        ):
            self.data["month"] = self.data["order_date"].dt.strftime("%Y-%m")

            if "total_amount" in self.data.columns:
                monthly_revenue = self.data.groupby("month")["total_amount"].sum()

                if not monthly_revenue.empty:
                    self.insights["monthly_revenue"] = monthly_revenue.to_dict()

                    if len(monthly_revenue) >= 2:
                        months = sorted(monthly_revenue.index)
                        current = monthly_revenue[months[-1]]
                        previous = monthly_revenue[months[-2]]

                        if previous > 0:
                            growth = ((current - previous) / previous) * 100
                            self.insights["monthly_growth"] = float(growth)

                            if growth < -10:
                                self.alerts.append(
                                    f"Sales declined by {abs(growth):.1f}% compared to previous month"
                                )
                            elif growth > 20:
                                self.alerts.append(
                                    f"Sales increased by {growth:.1f}% - consider increasing stock"
                                )

        # Customer-wise analysis
        if "customer_id" in self.data.columns:
            customer_order_count = self.data["customer_id"].value_counts()
            repeat_customers = (customer_order_count > 1).sum()

            self.insights["total_customers"] = int(customer_order_count.count())
            self.insights["repeat_customers"] = int(repeat_customers)

            if self.insights["total_customers"] > 0:
                repeat_rate = (
                    self.insights["repeat_customers"] / self.insights["total_customers"]
                ) * 100
                self.insights["repeat_rate"] = float(repeat_rate)

                if self.insights["repeat_rate"] < 10:
                    self.alerts.append(
                        f"Low repeat customer rate: {repeat_rate:.1f}%. Consider loyalty program."
                    )

        # Category performance analysis
        if "category" in self.data.columns and "total_amount" in self.data.columns:
            category_revenue = (
                self.data.groupby("category")["total_amount"]
                .sum()
                .sort_values(ascending=False)
            )
            self.insights["category_performace"] = category_revenue.to_dict()

            if not category_revenue.empty:
                self.insights["best_category"] = category_revenue.index[0]
                self.insights["worst_category"] = category_revenue.index[-1]

        self.logger.info(
            f"Analysis complete: {len(self.insights)} insights, {len(self.alerts)} alerts"
        )
        return self.insights

    def _standardize_column_names(self):
        if self.data is not None:
            column_mapping = {
                "order id": "order_id",
                "order date": "order_date",
                "product name": "product_name",
                "product": "product_name",
                "unit price": "unit_price",
                "total price": "total_amount",
                "revenue": "total_amount",
                "customer id": "customer_id",
                "customer": "customer_id",
                "qty": "quantity",
                "amount": "total_amount",
            }

            self.data.columns = [str(col).strip().lower() for col in self.data.columns]

            self.data.columns = [
                column_mapping.get(col, col) for col in self.data.columns
            ]

    def validate(self):
        return super().validate()
