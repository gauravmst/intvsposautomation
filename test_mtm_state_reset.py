import io

import pandas as pd

import mtm


def test_new_upload_clears_previous_cached_state():
    mtm._DATA_CACHE.clear()
    mtm._FILTER_CACHE.clear()
    mtm._EXCEL_CACHE.clear()

    mtm._DATA_CACHE["stale"] = ("old.csv", b"old")
    mtm._FILTER_CACHE["stale"] = (pd.DataFrame([{"user_id": "999"}]), {})
    mtm._EXCEL_CACHE["stale"] = ("old.xlsx", b"old")

    client = mtm.app.test_client()
    csv_bytes = b"date,algo,alias,user_id,allocation,mtm_all,dte\n2024-01-01,1,Alice,100,100,10,0DTE\n"

    response = client.post(
        "/process",
        data={"file": (io.BytesIO(csv_bytes), "new.csv")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    assert len(mtm._DATA_CACHE) == 1
    assert len(mtm._FILTER_CACHE) == 0
    assert "stale" not in mtm._DATA_CACHE
    assert "stale" not in mtm._FILTER_CACHE
    assert "stale" not in mtm._EXCEL_CACHE
    assert mtm._EXCEL_CACHE
