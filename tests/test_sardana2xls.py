import pytest
import os
import json


class DatabaseMock:
    def __init__(self):
        test_path = os.path.dirname(os.path.realpath(__file__))
        with open("{}/tangodb.json".format(test_path), "r") as fp:
            self.data = json.load(fp)


def test_test():
    db = DatabaseMock()
    assert not db.data
