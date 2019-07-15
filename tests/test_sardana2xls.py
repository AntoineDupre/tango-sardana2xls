import pytest
import os
import json
from sardana2xls import sardana2xls
from mock import MagicMock


class DatabaseMock:
    def __init__(self):
        test_path = os.path.dirname(os.path.realpath(__file__))
        with open("{}/tangodb.json".format(test_path), "r") as fp:
            self.data = json.load(fp)

    def get_device_name(self, server, cls):
        server, instance = server.split("/")
        if cls == "*":
            devices = []
            for cls in self.data["servers"][server][instance].values():
                devices += cls.keys()
            return devices
        return list(self.data["servers"][server][instance][cls].keys())

    def get_device_property(self, name, prop):
        for _, instance in self.data["servers"].items():
            for _, classes in instance.items():
                for _, devices in classes.items():
                    for device, props in devices.items():
                        if name.lower() == device.lower():
                            if prop == "Angleconversion":
                                print(prop)
                            return {
                                prop: props["properties"].get(prop.lower(), [])
                            }

    def get_device_property_list(self, name, props):
        for _, instance in self.data["servers"].items():
            for _, classes in instance.items():
                for _, devices in classes.items():
                    for device, props in devices.items():
                        if name.lower() == device.lower():
                            return props["properties"].keys()

    def get_alias(self, name):
        for a, instance in self.data["servers"].items():
            for _, classes in instance.items():
                for _, devices in classes.items():
                    for device, props in devices.items():
                        if name.lower() == device.lower():
                            return props.get("alias", "")

    # TODO: Fix it
    get_alias_from_device = get_alias

    def get_class_for_device(self, name):
        for a, instance in self.data["servers"].items():
            for _, classes in instance.items():
                for cls, devices in classes.items():
                    for device in devices.keys():
                        if name.lower() == device.lower():
                            return cls

    def get_db_host(self):
        return "hello"

    def get_db_port(self):
        return 1234

    def _get_attribute_props(self, name):
        ret = []
        for a, instance in self.data["servers"].items():
            for _, classes in instance.items():
                for cls, devices in classes.items():
                    for device, props in devices.items():
                        if name.lower() == device.lower():
                            attr = props.get("attribute_properties", {})
                            for name, value in attr.items():
                                try:
                                    ret.append(
                                        "{}:{}".format(
                                            name, value["__value"][0]
                                        )
                                    )
                                # TODO: Fix position min_value  ...
                                except KeyError:
                                    pass
                            return ret
        return ret


@pytest.fixture
def mock_db():
    sardana2xls.tango.Database = DatabaseMock
    database = DatabaseMock()
    sardana2xls.get_motor_attributes = lambda x: database._get_attribute_props(
        x
    )


@pytest.fixture
def sheet():
    return MagicMock()


@pytest.fixture
def smap(mock_db):
    return sardana2xls.SardanaMap("B108A")


def test_iors(smap, sheet):
    smap.proceed_iors(smap.iors, sheet)


def test_motors(smap, sheet):
    smap.proceed_motors(smap.motors, sheet)


def test_pseudo(smap, sheet):
    smap.proceed_pseudos(smap.pseudos, sheet)


def test_controller(smap, sheet):
    smap.proceed_controllers(smap.controllers, sheet)


def test_pool(smap, sheet):
    smap.proceed_pool(smap.pool_name, sheet)


def test_ms(smap, sheet):
    smap.proceed_macroserver(smap.ms_name, sheet)


def test_global(smap, sheet):
    smap.proceed_global(smap.pool, sheet)
    smap.proceed_iors(smap.iors, sheet)


def test_channel(smap, sheet):
    smap.proceed_channel(smap.channels, sheet)


def test_measgrps(smap, sheet):
    smap.proceed_measgrps(smap.measgrps, sheet)


def test_instruments(smap, sheet):
    smap.proceed_instruments(smap.instrument_list, sheet)


def test_doors(smap, sheet):
    smap.proceed_doors(smap.doors, sheet)
