from datetime import date, datetime

import pytest

from sheetflow.time_period import Period, generate_period_range


def test_generate_period_range_months_aligns_to_natural_months():
    periods = list(
        generate_period_range(
            {"start": date(2024, 1, 15), "end": date(2024, 2, 20), "freq": "M"}
        )
    )

    assert periods == [
        Period(date(2024, 1, 1), date(2024, 1, 31)),
        Period(date(2024, 2, 1), date(2024, 2, 29)),
    ]


def test_generate_period_range_months_cross_year():
    periods = list(
        generate_period_range(
            {"start": date(2024, 12, 15), "end": date(2025, 2, 10), "freq": "M"}
        )
    )

    assert periods == [
        Period(date(2024, 12, 1), date(2024, 12, 31)),
        Period(date(2025, 1, 1), date(2025, 1, 31)),
        Period(date(2025, 2, 1), date(2025, 2, 28)),
    ]


def test_generate_period_range_no_frequency_returns_original_span():
    periods = list(
        generate_period_range(
            {"start": date(2024, 1, 15), "end": date(2024, 2, 20), "freq": "N"}
        )
    )

    assert periods == [Period(date(2024, 1, 15), date(2024, 2, 20))]


def test_generate_period_range_no_frequency_normalizes_reversed_dates():
    periods = list(
        generate_period_range(
            {"start": date(2024, 2, 20), "end": date(2024, 1, 15), "freq": "N"}
        )
    )

    assert periods == [Period(date(2024, 1, 15), date(2024, 2, 20))]


def test_generate_period_range_quarters_aligns_to_natural_quarters():
    periods = list(
        generate_period_range(
            {"start": date(2024, 2, 15), "end": date(2024, 8, 20), "freq": "Q"}
        )
    )

    assert periods == [
        Period(date(2024, 1, 1), date(2024, 3, 31)),
        Period(date(2024, 4, 1), date(2024, 6, 30)),
        Period(date(2024, 7, 1), date(2024, 9, 30)),
    ]


def test_generate_period_range_years_aligns_to_natural_years():
    periods = list(
        generate_period_range(
            {"start": date(2023, 5, 15), "end": date(2024, 2, 20), "freq": "Y"}
        )
    )

    assert periods == [
        Period(date(2023, 1, 1), date(2023, 12, 31)),
        Period(date(2024, 1, 1), date(2024, 12, 31)),
    ]


def test_generate_period_range_years_multi_year():
    periods = list(
        generate_period_range(
            {"start": date(2022, 6, 1), "end": date(2025, 4, 30), "freq": "Y"}
        )
    )

    assert periods == [
        Period(date(2022, 1, 1), date(2022, 12, 31)),
        Period(date(2023, 1, 1), date(2023, 12, 31)),
        Period(date(2024, 1, 1), date(2024, 12, 31)),
        Period(date(2025, 1, 1), date(2025, 12, 31)),
    ]


def test_generate_period_range_reversed_dates_still_splits_forward():
    periods = list(
        generate_period_range(
            {"start": date(2024, 2, 20), "end": date(2024, 1, 15), "freq": "M"}
        )
    )

    assert periods == [
        Period(date(2024, 1, 1), date(2024, 1, 31)),
        Period(date(2024, 2, 1), date(2024, 2, 29)),
    ]


def test_same_month_returns_single_period():
    periods = list(
        generate_period_range(
            {
                "start": date(2024, 1, 15),
                "end": date(2024, 1, 20),
                "freq": "M",
            }
        )
    )

    assert periods == [
        Period(date(2024, 1, 1), date(2024, 1, 31)),
    ]


def test_same_day_returns_single_day():
    periods = list(
        generate_period_range(
            {"start": date(2024, 6, 15), "end": date(2024, 6, 15), "freq": "M"}
        )
    )

    assert periods == [Period(date(2024, 6, 15), date(2024, 6, 15))]


def test_datetime_and_tuple_inputs_are_normalized_to_dates():
    periods = list(
        generate_period_range(
            {
                "start": datetime(2024, 1, 15, 8, 30),
                "end": (2024, 2, 20),
                "freq": "N",
            }
        )
    )

    assert periods == [Period(date(2024, 1, 15), date(2024, 2, 20))]


def test_invalid_freq_raises_at_call_time():
    with pytest.raises(ValueError, match="Invalid freq"):
        generate_period_range(
            {"start": date(2024, 1, 1), "end": date(2024, 6, 30), "freq": "X"}
        )


def test_invalid_freq_is_rejected_for_same_day():
    with pytest.raises(ValueError, match="Invalid freq"):
        generate_period_range(
            {"start": date(2024, 1, 1), "end": date(2024, 1, 1), "freq": "X"}
        )


def test_invalid_date_type_raises_at_call_time():
    with pytest.raises(TypeError, match="Date value must be"):
        generate_period_range(
            {"start": "2024-01-01", "end": date(2024, 1, 31), "freq": "M"}
        )


@pytest.mark.parametrize("freq", ["M", "Q", "Y"])
def test_natural_periods_support_maximum_date(freq):
    periods = list(
        generate_period_range(
            {"start": date(9999, 10, 15), "end": date.max, "freq": freq}
        )
    )

    assert periods[-1].end == date.max
