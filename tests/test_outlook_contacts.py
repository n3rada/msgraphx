# tests/test_outlook_contacts.py

from __future__ import annotations

from collections import Counter

from msgraphx.modules.outlook.contacts import _rank_contacts


def test_rank_contacts_orders_by_to_then_cc_then_address():
    to_counter = Counter({"bob@corp.com": 5, "alice@corp.com": 2, "carol@corp.com": 2})
    cc_counter = Counter({"carol@corp.com": 4, "dave@corp.com": 9})
    names = {"bob@corp.com": "Bob"}

    ranked = _rank_contacts(to_counter, cc_counter, names)

    assert [r["email"] for r in ranked] == [
        "bob@corp.com",
        "carol@corp.com",  # ties alice on To, wins on CC
        "alice@corp.com",
        "dave@corp.com",  # CC only
    ]
    assert [r["rank"] for r in ranked] == [1, 2, 3, 4]
    assert ranked[0] == {
        "rank": 1,
        "email": "bob@corp.com",
        "name": "Bob",
        "to_count": 5,
        "cc_count": 0,
    }
    assert ranked[3]["name"] is None
    assert ranked[3]["to_count"] == 0 and ranked[3]["cc_count"] == 9


def test_rank_contacts_top_limits_results():
    to_counter = Counter({"a@x.io": 3, "b@x.io": 2, "c@x.io": 1})

    ranked = _rank_contacts(to_counter, Counter(), {}, top=2)

    assert [r["email"] for r in ranked] == ["a@x.io", "b@x.io"]


def test_rank_contacts_top_none_keeps_all():
    to_counter = Counter({"a@x.io": 1, "b@x.io": 1})

    assert len(_rank_contacts(to_counter, Counter(), {}, top=None)) == 2


def test_rank_contacts_empty():
    assert _rank_contacts(Counter(), Counter(), {}) == []
