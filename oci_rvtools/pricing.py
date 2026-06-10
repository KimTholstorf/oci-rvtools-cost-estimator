# Copyright (c) 2026 Kim Tholstorf
# https://github.com/KimTholstorf/oci-rvtools-cost-estimator
# MIT License — see LICENSE file for details

"""Client for the Oracle CETools list-price API."""

from __future__ import annotations

import json
from typing import Dict, Optional
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

from .model import PriceRecord

API_BASE = "https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/"


class PricingClient:
    def __init__(self, currency: str) -> None:
        self.currency = currency.upper()
        self._cache: Dict[str, PriceRecord] = {}

    def get_price(self, part_number: str) -> PriceRecord:
        part_number = part_number.strip()
        if part_number in self._cache:
            return self._cache[part_number]

        params = {"partNumber": part_number, "currencyCode": self.currency}
        url = f"{API_BASE}?{urlparse.urlencode(params)}"
        try:
            with urlrequest.urlopen(url) as response:
                payload = response.read().decode("utf-8")
        except urlerror.URLError as exc:
            raise RuntimeError(f"Failed to reach OCI price API ({url}): {exc}") from exc

        try:
            data = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Invalid JSON from OCI price API for part {part_number}") from exc

        items = data.get("items")
        if not isinstance(items, list) or not items:
            raise RuntimeError(f"No pricing data for part {part_number} ({self.currency})")

        price: Optional[float] = None
        display_name: Optional[str] = None
        for item in items:
            if not isinstance(item, dict):
                continue
            if item.get("partNumber") and item["partNumber"] != part_number:
                continue
            price = self._extract_price(item, self.currency)
            display_name = item.get("displayName") or display_name
            if price is not None:
                break

        if price is None:
            for item in items:
                if isinstance(item, dict):
                    price = PricingClient._extract_price(item, self.currency)
                    display_name = item.get("displayName") or display_name
                    if price is not None:
                        break

        if price is None:
            raise RuntimeError(f"Could not determine unit price for part {part_number}: {payload[:500]}...")

        if not display_name:
            display_name = part_number

        record = PriceRecord(part_number=part_number, display_name=str(display_name), unit_price=float(price))
        self._cache[part_number] = record
        return record

    @staticmethod
    def _extract_price(item: Dict[str, object], currency: str) -> Optional[float]:
        currency = currency.upper()
        candidate_keys = [
            "price",
            "unitPrice",
            "unit_price",
            "unit_price_value",
            "unit_price_included",
            "netUnitPrice",
            "list_price",
            "usdPrice",
            "amount",
        ]
        for key in candidate_keys:
            if key in item:
                try:
                    return float(item[key])  # type: ignore[arg-type]
                except (TypeError, ValueError):
                    continue

        if "prices" in item and isinstance(item["prices"], list):
            for entry in item["prices"]:
                if isinstance(entry, dict):
                    maybe = PricingClient._extract_price(entry, currency)
                    if maybe is not None:
                        return maybe

        locs = item.get("currencyCodeLocalizations")
        if isinstance(locs, list):
            for loc in locs:
                if not isinstance(loc, dict):
                    continue
                loc_currency = loc.get("currencyCode")
                if loc_currency and loc_currency.upper() != currency:
                    continue
                prices = loc.get("prices")
                if isinstance(prices, list):
                    for entry in prices:
                        if not isinstance(entry, dict):
                            continue
                        if "value" in entry:
                            try:
                                return float(entry["value"])
                            except (TypeError, ValueError):
                                pass
                        maybe = PricingClient._extract_price(entry, currency)
                        if maybe is not None:
                            return maybe
        return None
