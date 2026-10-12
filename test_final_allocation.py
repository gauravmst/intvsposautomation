import datetime as dt

import pandas as pd

import mtm


def _df(rows):
    df = pd.DataFrame(rows, columns=["date", "algo", "user_id", "allocation", "mtm_all", "dte"])
    df["date"] = pd.to_datetime(df["date"])
    return df


def _alloc(out, user, date):
    return out[(out["user_id"] == user) & (out["date"] == date)]["allocation"].iloc[0]


def test_legs_take_zero_dte_allocation_per_cycle():
    df = _df([
        # NF cycle: Fri 4DTE, Mon 1DTE, Tue 0DTE -> expiry Tue 2026-09-08
        ("2026-09-04", 1, "A", 100, 10, "4DTE"),
        ("2026-09-07", 1, "A", 150, 10, "1DTE"),
        ("2026-09-08", 1, "A", 200, 10, "0DTE"),
        # SX cycle: Wed 1DTE, Thu 0DTE -> expiry Thu 2026-09-10
        ("2026-09-09", 1, "A", 120, 10, "1DTE"),
        ("2026-09-10", 1, "A", 180, 10, "0DTE"),
        # B is stopped out on Wed 1DTE (no 0DTE) and must not borrow NF's 0DTE
        ("2026-09-09", 1, "B", 90, -5, "1DTE"),
        ("2026-09-11", 1, "B", 70, 3, "4DTE"),
        ("2026-09-15", 1, "B", 110, 3, "0DTE"),
        # C is not in the positional file -> untouched
        ("2026-09-07", 1, "C", 50, 1, "1DTE"),
        ("2026-09-08", 1, "C", 80, 1, "0DTE"),
    ])
    d = dt.date
    pos = {"A": {d(2026, 9, 4), d(2026, 9, 7), d(2026, 9, 9)},
           "B": {d(2026, 9, 9), d(2026, 9, 11)}}

    out, stats = mtm.apply_final_allocation(df, pos)

    ts = pd.Timestamp
    assert [_alloc(out, "A", ts(x)) for x in ("2026-09-04", "2026-09-07", "2026-09-08")] == [200] * 3
    assert [_alloc(out, "A", ts(x)) for x in ("2026-09-09", "2026-09-10")] == [180] * 2
    assert _alloc(out, "B", ts("2026-09-09")) == 90          # stopped out: last leg
    assert _alloc(out, "B", ts("2026-09-11")) == 110         # 4DTE -> Tue 15th 0DTE
    assert _alloc(out, "C", ts("2026-09-07")) == 50
    assert stats["no_zero_dte"] == 1
    assert df["allocation"].tolist()[0] == 100               # input not mutated


def test_no_positional_file_is_a_noop():
    df = _df([("2026-09-08", 1, "A", 200, 10, "0DTE")])
    out, stats = mtm.apply_final_allocation(df, None)
    assert out is df and stats["cycles"] == 0


def test_month_end_cycle_gets_label_and_stays_pos_next_month():
    df = _df([
        ("2026-09-30", 8, "A", 200, -60, "1DTE"),
        ("2026-10-01", 8, "A", 300, -75, "0DTE"),
        ("2026-10-01", 8, "I", 50, 5, "0DTE"),          # plain intraday user
    ])
    pos = {"A": {dt.date(2026, 9, 30)}}

    out, _ = mtm.apply_final_allocation(df, pos)
    assert out["cycle"].tolist() == ["300926-011026", "300926-011026", "011026-011026"]
    assert out["allocation"].tolist()[:2] == [300, 300]

    oct_rows = out[out["date"] >= "2026-10-01"]
    ru = mtm.build_account_rollup(oct_rows, pos).set_index("UserID")
    assert ru.loc["A", "Type"] == "POS"      # 0DTE row alone in October, still positional
    assert ru.loc["I", "Type"] == "INT"
