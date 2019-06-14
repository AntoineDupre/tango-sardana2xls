from sardana2xls import utils
import pytest


def test_unique_dict():
    dic = utils.unique_dict()
    dic["Test"] = 1
    assert "Test" in dic
    assert dic["Test"] == 1
    dic["Test2"] = 2
    assert "Test2" in dic
    assert dic["Test2"] == 2
    dic["Test2"] = 1
    assert "Test2" in dic
    assert "Test" not in dic
    assert dic["Test2"] == 1


def test_unique_bidict():
    dic = utils.unique_bidict()
    dic["Test"] = 1
    assert "Test" in dic
    assert 1 in dic
    assert dic["Test"] == 1
    assert dic[1] == "Test"
    dic["Test2"] = 2
    assert "Test2" in dic
    assert 2 in dic
    assert dic["Test2"] == 2
    assert dic[2] == "Test2"
    dic["Test2"] = 1
    assert "Test2" in dic
    assert "Test" not in dic
    assert dic["Test2"] == 1
    assert dic[1] == "Test2"
