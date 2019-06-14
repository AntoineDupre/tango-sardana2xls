from sardana2xls.utils import get_elements, get_ms_elements
from sardana2xls.utils import generate_aliases_mapping
from sardana2xls.utils import generate_id_mapping
from sardana2xls.utils import generate_prop_mapping
from sardana2xls.utils import generate_class_mapping
from sardana2xls.utils import generate_instrument_list
from sardana2xls.utils import generate_instrument_mapping
import tango
import xlrd
from xlutils.copy import copy
from functools import partial
import sys
import os
import argparse
import logging


class SardanaMap:
    """ Manage sardana elements """
    def __init__(self, pool):
        # Connect to the tangodb
        self.db = tango.Database()
        # Collect tango device running in the Pool
        self.pool = pool
        self.pool_server = "Pool/{}".format(self.pool)
        self.pool_name = self.db.get_device_name(self.pool_server, "Pool")[0]
        logging.info("Pool: {}".format(pool))
        logging.info("Server: {}".format(self.pool_server))
        logging.info("Pool device: {}".format(self.pool_name))
        # Collect tango device running in the MS
        self.ms_server = "MacroServer/{}".format(pool)
        self.ms_name = self.db.get_device_name(self.ms_server, "MacroServer")[
            0
        ]
        logging.info("MacroServer: {}".format(self.ms_server))
        logging.info("MacroServer device: {}".format(self.ms_name))
        self._setup_mapping()
        self._setup_class_mapping()

    def _setup_mapping(self):
        """ Generate internal sardana id mapping """
        db = self.db
        # Prepare environment
        elements = get_elements(self.pool, self.db)
        self.elements = elements
        self.ms_elements = get_ms_elements(self.pool, self.db)
        # Generate mapping
        self.aliases = generate_aliases_mapping(elements, db)
        self.ids = generate_id_mapping(elements, db)
        self.ctrl_ids = generate_prop_mapping(elements, db, "ctrl_id")
        self.motor_ids = generate_prop_mapping(elements, db, "motor_role_ids")
        self.pseudo_ids = generate_prop_mapping(
            elements, db, "pseudo_motor_role_ids"
        )
        self.channel_ids = generate_prop_mapping(elements, db, "elements")
        self.instrument_list = generate_instrument_list(self.pool_name, db)
        self.instrument_ids = generate_instrument_mapping(self.instrument_list)

    def _setup_class_mapping(self):
        """ Sort sardana elements  """
        elements = self.elements
        ms_elements = self.ms_elements
        db = self.db
        # Class mapping
        classes = generate_class_mapping(elements, db)
        self.classes = classes
        classes_ms = generate_class_mapping(ms_elements, db)
        self.classes_ms = classes_ms
        self.controllers = [
            k for k, v in classes.items() if v.lower() == "controller"
        ]
        self.motors = [k for k, v in classes.items() if v == "Motor"]
        self.pseudos = [k for k, v in classes.items() if v == "PseudoMotor"]
        self.iors = [k for k, v in classes.items() if v == "IORegister"]
        self.measgrps = [
            k for k, v in classes.items() if v == "MeasurementGroup"
        ]
        self.macroservers = [
            k for k, v in classes_ms.items() if v == "MacroServer"
        ]
        self.doors = [k for k, v in classes_ms.items() if v == "Door"]
        self.channels = [
            (k, v)
            for k, v in classes.items()
            if "counter" in v.lower() or "channel" in v.lower()
        ]

    def ior_data(self, name):
        """ Format one IORegister """
        ior_ctrl = self.aliases[self.ids[self.ctrl_ids[name][0]]]
        ior_type = "IORegister"
        ior_pool = self.pool_name
        ior_alias = self.aliases[name]
        ior_name = name
        ior_axis = get_property(name, "Axis")
        try:
            ior_instr = get_property(name, "instrument_id")
            ior_instrument = self.instrument_ids[ior_instr]
        # TODO: Which exception ?
        except Exception:
            ior_instrument = ""
        ior_desc = ""
        ior_attributes = ";".join(get_motor_attributes(name))
        return (
            ior_type,
            ior_pool,
            ior_ctrl,
            ior_alias,
            ior_name,
            ior_axis,
            ior_instrument,
            ior_desc,
            ior_attributes,
        )

    def channel_data(self, name, _type):
        """ Format one Acquisition channel """
        channel_ctrl = self.aliases[self.ids[self.ctrl_ids[name][0]]]
        channel_type = _type
        channel_pool = self.pool_name
        channel_alias = self.aliases[name]
        channel_name = name
        channel_axis = get_property(name, "Axis")
        try:
            channel_instr = get_property(name, "instrument_id")
            channel_instrument = self.instrument_ids[channel_instr]
        # TODO: Which exception ?
        except Exception:
            channel_instrument = ""
        channel_desc = ""
        channel_attributes = ";".join(get_motor_attributes(name))
        return (
            channel_type,
            channel_pool,
            channel_ctrl,
            channel_alias,
            channel_name,
            channel_axis,
            channel_instrument,
            channel_desc,
            channel_attributes,
        )

    def motor_data(self, name, mot_type):
        """ Format one motor """
        mot_type = mot_type
        mot_pool = self.pool_name

        mot_ctrl = self.aliases[self.ids[self.ctrl_ids[name][0]]]
        try:
            mot_alias = self.aliases[name]
        except KeyError:
            mot_alias = ""
        mot_device = name
        mot_axis = get_property(name, "Axis")
        try:
            mot_instr = get_property(name, "instrument_id")
            mot_instrument = self.instrument_ids[mot_instr]
        # TODO: Which exception
        except Exception:
            mot_instrument = ""
        mot_desc = ""
        mot_attributes = ";".join(get_motor_attributes(name))
        return (
            mot_type,
            mot_pool,
            mot_ctrl,
            mot_alias,
            mot_device,
            mot_axis,
            mot_instrument,
            mot_desc,
            mot_attributes,
        )

    def get_controller_elements(self, name, ctrl_type):
        elems = []
        if ctrl_type == "PseudoMotor":
            for motor in self.motor_ids[name]:
                try:
                    elems.append(self.aliases[self.ids[motor]])
                except KeyError as e:
                    print(e)
        return ";".join(elems)

    def controller_data(self, name):
        ctrl_prop = partial(get_property, name)
        ctrl_type = ctrl_prop("type")
        ctrl_lib = ctrl_prop("library")
        ctrl_class = ctrl_prop("klass")
        ctrl_props = ";".join(get_properties(name))
        ctrl_elements = self.get_controller_elements(name, ctrl_type)
        # ctrl_device = name
        return [
            ctrl_type,
            self.pool_name,
            self.aliases[name],
            # ctrl_device,
            ctrl_lib,
            ctrl_class,
            ctrl_props,
            ctrl_elements,
        ]

    def mg_data(self, name):
        mg_type = "MeasurementGroup"
        mg_pool = self.pool_name
        mg_device = name
        mg_alias = self.aliases[name]
        mg_desc = ""
        mg_channels = self.get_mg_channels(name)

        return (mg_type, mg_pool, mg_alias, mg_device, mg_channels, mg_desc)

    def get_mg_channels(self, name):
        elems = []
        for chan in self.channel_ids[name]:
            try:
                elems.append(self.aliases[self.ids[chan]])
            except KeyError as e:
                print(e)
        return ";".join(elems)

    def proceed_controllers(self, names, sheet):
        logging.info("Create controllers")
        ctrls = []
        for ctrl in names:
            logging.info("{}".format(ctrl))
            data = self.controller_data(ctrl)
            ctrls.append(data)
        ctrls = sorted(ctrls, key=lambda x: (x[0], x[2]))
        for line, data in enumerate(ctrls):
            write_line(sheet, line + 1, data)

    def proceed_motors(self, names, sheet):
        logging.info("Create motors")
        motors = []
        for motor in names:
            data = self.motor_data(motor, "Motor")
            motors.append(data)
        motors = sorted(motors, key=lambda x: (x[2], int(x[5])))
        for line, data in enumerate(motors):
            write_line(sheet, line + 1, data)

    def proceed_pseudos(self, names, sheet):
        logging.info("Create pseudo motors")
        pseudos = []
        for motor in names:
            data = self.motor_data(motor, "PseudoMotor")
            pseudos.append(data)
        pseudos = sorted(pseudos, key=lambda x: (x[2], int(x[5])))
        for line, data in enumerate(pseudos):
            write_line(sheet, line + 1, data)

    def proceed_pool(self, name, sheet):
        # get_properties
        host = ":".join((self.db.get_db_host(), str(self.db.get_db_port())))
        pool_alias = self.db.get_alias_from_device(self.pool_name)
        prop = "\n".join(
            self.db.get_device_property(self.pool_name, "PoolPath")["PoolPath"]
        )
        line = (
            "Pool",
            host,
            self.pool_server,
            "",  # Description
            pool_alias,  # Alias
            self.pool_name,
            prop,
        )
        write_line(sheet, 1, line)

    def proceed_macroserver(self, name, sheet):
        # get_properties
        host = ":".join((self.db.get_db_host(), str(self.db.get_db_port())))
        ms_alias = self.db.get_alias_from_device(self.ms_name)
        prop = "\n".join(
            self.db.get_device_property(self.ms_name, "MacroPath")["MacroPath"]
        )
        pools = "\n".join(
            self.db.get_device_property(self.ms_name, "PoolNames")["PoolNames"]
        )
        line = (
            "MacroServer",
            host,
            self.ms_server,
            "",  # Description
            ms_alias,  # Alias
            self.ms_name,
            prop,
            pools,
        )
        write_line(sheet, 2, line)

    def proceed_doors(self, names, sheet):
        # Server	MacroServer	Description	name	tango name
        for line, name in enumerate(names):
            data = (
                self.ms_server,
                self.ms_name,
                "enter description",  # Sardana dsconfig doesnt like empty description
                self.db.get_alias_from_device(name),
                name,
            )
            write_line(sheet, line + 1, data)

    def proceed_global(self, name, sheet):
        write_line(sheet, 0, ("code", self.pool))
        write_line(sheet, 1, ("name", self.pool))
        write_line(sheet, 2, ("description",))
        write_line(sheet, 3, ("",))
        write_line(sheet, 4, ("prefix", "p1"))

    def proceed_measgrps(self, names, sheet):
        logging.info("Create measurement groups")
        _mgs = []
        for mg in names:
            data = self.mg_data(mg)
            _mgs.append(data)
        _mgs = sorted(_mgs, key=lambda x: (x[2], x[5]))
        for line, data in enumerate(_mgs):
            write_line(sheet, line + 1, data)

    def proceed_instruments(self, instr_list, sheet):
        logging.info("Create instruments")
        for line, data in enumerate(instr_list):
            instr_type = "Instrument"
            instr_pool = self.pool_name
            instr_name = data[1]
            instr_class = data[0]
            line_data = (instr_type, instr_pool, instr_name, instr_class)
            write_line(sheet, line + 1, line_data)

    def proceed_iors(self, names, sheet):
        logging.info("Create ioregister")
        _iors = []
        for ior in names:
            data = self.ior_data(ior)
            _iors.append(data)
        _iors = sorted(_iors, key=lambda x: (x[2], x[5]))
        for line, data in enumerate(_iors):
            write_line(sheet, line + 1, data)

    def proceed_channel(self, names, sheet):
        logging.info("Create channels")
        _channels = []
        for channel, _type in names:
            data = self.channel_data(channel, _type)
            _channels.append(data)
        _channels = sorted(_channels, key=lambda x: (x[2], x[5]))
        for line, data in enumerate(_channels):
            write_line(sheet, line + 1, data)


class XlsWriter:
    def __init__(self, template):
        # Open xls file
        module_path = os.path.dirname(os.path.realpath(__file__))
        template_path = "{}/{}".format(module_path, template)
        r_workbook = xlrd.open_workbook(template_path)
        w_workbook = copy(r_workbook)
        self.w_workbook = w_workbook
        self.door_sheet = w_workbook.get_sheet(2)
        self.controller_sheet = w_workbook.get_sheet(3)
        self.motor_sheet = w_workbook.get_sheet(4)
        self.pseudo_sheet = w_workbook.get_sheet(5)
        self.servers_sheet = w_workbook.get_sheet(1)
        self.global_sheet = w_workbook.get_sheet(0)
        self.ior_sheet = w_workbook.get_sheet(6)
        self.channel_sheet = w_workbook.get_sheet(7)
        self.measurment_sheet = w_workbook.get_sheet(8)
        self.acq_sheet = w_workbook.get_sheet(9)
        self.instr_sheet = w_workbook.get_sheet(11)


default_properties = [
    "id",
    "ctrl_id",
    "motor_role_ids",
    "pseudo_motor_role_ids",
    "type",
    "library",
    "klass",
    "__SubDevices",
]

mot_attributes = [
    "EncoderSource",
    "EncoderSourceFormula",
    "Sign",
    "Offset",
    "Step_per_unit",
    "UserEncoderSource",
]


def get_property(ds, name):
    db = tango.Database()
    proplist = db.get_device_property(ds, name)[name]
    if len(proplist) > 1:
        prop = "\\n".join(proplist)
    else:
        prop = proplist[0]
    return prop


def get_property_list(name):
    db = tango.Database()
    return [
        p
        for p in db.get_device_property_list(name, "*")
        if p not in default_properties
    ]


def get_properties(name):
    props = get_property_list(name)
    return ["{}:{}".format(p, get_property(name, p)) for p in props]


def write_line(sheet, line, data):
    for index, d in enumerate(data):
        sheet.write(line, index, d)


def get_motor_attributes(name):
    tango_db = tango.DeviceProxy("sys/database/2")
    quer = "Select attribute, value from property_attribute_device "
    query += "where device='{}' and name='__value'"
    query = query.format(name)
    reply = tango_db.DbMySqlSelect(query)
    reply = reply[1]
    if "DialPosition" in reply:
        idx = reply.index("DialPosition")
        del reply[idx : idx + 2]
    if "PowerOn" in reply:
        idx = reply.index("PowerOn")
        del reply[idx : idx + 2]
    answer = [
        "{}:{}".format(att, value)
        for att, value in zip(reply[::2], reply[1::2])
    ]
    return answer


def proceed(pool_name):
    smap = SardanaMap(pool_name)
    writer = XlsWriter("template/template.xls")
    smap.proceed_motors(smap.motors, writer.motor_sheet)
    smap.proceed_pseudos(smap.pseudos, writer.pseudo_sheet)
    smap.proceed_controllers(smap.controllers, writer.controller_sheet)
    smap.proceed_pool(smap.pool_name, writer.servers_sheet)
    smap.proceed_macroserver(smap.ms_name, writer.servers_sheet)
    smap.proceed_global(smap.pool, writer.global_sheet)
    smap.proceed_iors(smap.iors, writer.ior_sheet)
    smap.proceed_channel(smap.channels, writer.channel_sheet)
    smap.proceed_measgrps(smap.measgrps, writer.acq_sheet)
    smap.proceed_instruments(smap.instrument_list, writer.instr_sheet)
    smap.proceed_doors(smap.doors, writer.door_sheet)

    writer.w_workbook.save("{}/{}.xls".format(os.getcwd(), pool_name))


def main():
    logging.basicConfig(level=logging.DEBUG)
    usage = "%prog [options] <pool_instance> "
    parser = argparse.ArgumentParser(usage)
    parser.add_argument('poolname', metavar='pool', type=str,
                    help='Pool instance name')
    args = parser.parse_args()
    proceed(args.poolname)


if __name__ == "__main__":
    main()
