import pandas as pd
from .base_processor import BaseProcessor
from utils.formatters import format_currency


class InventoryProcessor(BaseProcessor):

    def load_data(self, filepath):

        self.logger.info(f"Loading inventory data from {filepath}")

        try:
            self.data = pd.read_csv(filepath)
            self._standardize_columns_names()

            numeric_column = [
                "current_stock",
                "reorder_level",
                "unit_cost",
                "days_of_supply",
            ]
            for col in numeric_column:
                if col in self.data.columns:
                    self.data[col] = pd.to_numeric(
                        self.data[col], errors="coerce"
                    ).fillna(0)

            self.logger.info(f"Loded {len(self.data)} inventory items")
            return self

        except Exception as e:
            self.logger.error(f"Error loading inventory data: {e}")
            raise

    def analyze(self):
        self.logger.info(f"Analyzing inventory data...")
        self.validate()

        self.insights = {}
        self.alerts = []

        required_cols = ["product_name", "current_stock"]

        missing = [col for col in required_cols if col not in self.data.columns]

        if missing:
            raise ValueError("Missing required files: {missing}")

        # Basic metrics
        self.insights["total_products"] = int(len(self.data))
        self.insights["total_stock_units"] = int(self.data["current_stock"].sum())

        if "unit_cost" in self.data.columns:
            self.data["inventory_value"] = (
                self.data["unit_cost"] * self.data["current_stock"]
            )
            self.insights["total_inventory_value"] = float(
                self.data["inventory_value"].sum()
            )

        if "reorder_level" in self.data.columns:
            low_stock_mask = self.data["current_stock"] <= self.data["reorder_level"]
            low_stock_items = self.data[low_stock_mask]

            self.insights["low_stock_count"] = int(len(low_stock_items))
            self.insights["low_stock_items"] = low_stock_items[
                ["product_name", "current_stock", "reorder_level"]
            ].to_dict("records")

            if not low_stock_items.empty:
                for _, item in low_stock_items.iterrows():
                    product = item["product_name"]
                    current = item["current_stock"]
                    reorder = item["reorder_level"]

                if current == 0:
                    self.alerts.append(f"CRITICAL: {product} is OUT OF STOCK")
                elif current < reorder * 0.5:
                    self.alerts.append(
                        f"URGENT: {product} has only {current} units (reorder: {reorder})"
                    )
                else:
                    self.alerts.append(
                        f"WARNING: {product} needs restock ({current}/{reorder})"
                    )

        # Identify items with excessive stock (more than 3x reorder level or >90 days supply)
        if "reorder_level" in self.data.columns:
            over_stock_mask = (self.data["reorder_level"] * 3) < self.data[
                "current_stock"
            ]
            over_stock_items = self.data[over_stock_mask]

            self.insights["overstock_count"] = int(len(over_stock_items))
            self.insights["overstock_items"] = over_stock_items[
                ["product_name", "current_stock", "reorder_level"]
            ].to_dict("records")

            if not over_stock_items.empty:
                total_overstock_value = 0
                if "inventory" in self.data.columns:
                    total_overstock_value = over_stock_items["inventory_value"].sum()

                    self.alerts.append(
                        f"{len(over_stock_items)} products are overstocked "
                        f"(tying up {format_currency(total_overstock_value)} in capital)"
                    )

        # Day of supply analysis
        if "days_of_supply" in self.data.columns:
            critical_mask = self.data["days_of_supply"] < 3
            critical_items = self.data[critical_mask]

            low_mask = self.data["days_of_supply"] < 7
            low_items = self.data[low_mask]

            excessive_mask = self.data["days_of_supply"] > 90
            excessive_items = self.data[excessive_mask]

            self.insights["critical_items"] = int(len(critical_items))
            self.insights["low_items"] = int(len(low_items))
            self.insights["excessive_items"] = int(len(excessive_items))

            if not critical_items.empty:
                critical_names = ", ".join(
                    critical_items["product_name"].head(3).tolist()
                )
                self.alerts.append(
                    f"CRITICAL: {len(critical_items)} items have <3 days supply: {critical_names}"
                )

        # Category analysis
        if "category" in self.data.columns:
            category_stock = self.data.groupby("category").agg(
                {"current_stock": "sum", "product_name": "count"}
            )

            category_stock.columns = ["total_stock", "product_name"]

            for category, data in category_stock.iterrows():
                if data["total_stock"] == 0:
                    self.alerts.append(f"Category '{category}' has zero stock")

        # Restock recommendations
        if "reorder_level" in self.data.columns and "unit_cost" in self.data.columns:
            restock_recommendations = []
            for _, item in self.data.iterrows():
                current = item["current_stock"]
                reorder = item["reorder_level"]

                if current < reorder:
                    needed = reorder * 2 - current
                    cost = needed * item["unit_cost"]

                    if needed > 0:
                        restock_recommendations.append(
                            {
                                "product": item["product_name"],
                                "order_quantity": int(needed),
                                "estimated_cost": float(cost),
                                "urgency": "HIGH" if current == 0 else "MEDIUM",
                            }
                        )

        self.insights["restock_recommendations"] = restock_recommendations

        # Calculate total restock cost
        if restock_recommendations:
            total_cost = sum(rec["estimated_cost"] for rec in restock_recommendations)
            self.insights["total_restock_cost"] = float(total_cost)
            self.alerts.append(
                f"Restocking needed: {format_currency(total_cost)} for {len(restock_recommendations)} products"
            )

        self.logger.info(
            f"Inventory analysis complete: {len(self.insights)} insights, {len(self.alerts)} alerts"
        )

        return self

    def _standardize_column_names(self):
        """Standardize inventory specific column names."""
        if self.data is not None:
            column_mapping = {
                "product name": "product_name",
                "product": "product_name",
                "stock": "current_stock",
                "available_stock": "current_stock",
                "quantity": "current_stock",
                "reorder point": "reorder_level",
                "min_stock": "reorder_level",
                "cost": "unit_cost",
                "price": "unit_cost",
                "days supply": "days_of_supply",
                "supply_days": "days_of_supply",
            }

            self.data.columns = [
                str(col).strip().lower().replace(" ", "_") for col in self.data.columns
            ]
            self.data.columns = [
                column_mapping.get(col, col) for col in self.data.columns
            ]

    def validate(self):
        return super().validate()
