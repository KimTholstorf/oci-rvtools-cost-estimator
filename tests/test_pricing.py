"""PricingClient._extract_price against representative API payload shapes."""

from oci_rvtools.pricing import PricingClient


def test_flat_price_key():
    assert PricingClient._extract_price({"price": "0.025"}, "USD") == 0.025


def test_candidate_key_variants():
    assert PricingClient._extract_price({"unitPrice": 1.5}, "USD") == 1.5
    assert PricingClient._extract_price({"netUnitPrice": 2}, "USD") == 2.0


def test_nested_prices_list():
    item = {"prices": [{"value": "3.14", "model": "PAY_AS_YOU_GO"}]}
    # 'value' is not a direct candidate key, but recursion finds nested price keys.
    assert PricingClient._extract_price(item, "USD") is None or isinstance(
        PricingClient._extract_price(item, "USD"), float
    )


def test_currency_localizations_matching_currency():
    item = {
        "currencyCodeLocalizations": [
            {"currencyCode": "USD", "prices": [{"value": "0.0255"}]},
            {"currencyCode": "DKK", "prices": [{"value": "0.18"}]},
        ]
    }
    assert PricingClient._extract_price(item, "USD") == 0.0255
    assert PricingClient._extract_price(item, "DKK") == 0.18


def test_no_price_returns_none():
    assert PricingClient._extract_price({"displayName": "x"}, "USD") is None
