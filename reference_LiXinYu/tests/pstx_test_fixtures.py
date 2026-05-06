# -*- coding: utf-8 -*-
"""Shared fixtures for PSTX parser/page/rule tests."""

def make_cap(refdes='C1', rated='6.3V', bom_option='', page='PAGE1', page_real=''):
    display_page = page_real or ''
    return {
        'refdes': refdes,
        'part_name': 'CAP_0402',
        'hq_code': '',
        'value': '0.1uF',
        'package': '0402',
        'material': '',
        'tolerance': '',
        'voltage': rated,
        'current': '',
        'power': '',
        'bom_option': bom_option,
        'bom_cost': '',
        'room': '',
        'drawing': 'SCH_PAGE1',
        'page': display_page,
        'page_logical': page,
        'page_real': display_page,
        'comp_type': 'CAP',
        'nets': {},
    }

def make_ic(refdes='U1', nets=None, bom_option='', page='PAGE1', page_real=''):
    display_page = page_real or ''
    return {
        'refdes': refdes,
        'part_name': 'IC_CPU',
        'hq_code': 'PN_IC',
        'value': 'CPU',
        'package': 'BGA',
        'material': '',
        'tolerance': '',
        'voltage': '',
        'current': '',
        'power': '',
        'bom_option': bom_option,
        'bom_cost': '',
        'room': '',
        'drawing': 'SCH_PAGE1',
        'page': display_page,
        'page_logical': page,
        'page_real': display_page,
        'comp_type': 'IC',
        'nets': nets or {},
    }

def make_res(refdes, net_a, net_b, value='10k', bom_option='', page='PAGE1', page_real=''):
    display_page = page_real or ''
    return {
        'refdes': refdes,
        'part_name': 'RES_0402',
        'hq_code': '',
        'value': value,
        'package': '0402',
        'material': '',
        'tolerance': '',
        'voltage': '',
        'current': '',
        'power': '',
        'bom_option': bom_option,
        'bom_cost': '',
        'room': '',
        'drawing': 'SCH_PAGE1',
        'page': display_page,
        'page_logical': page,
        'page_real': display_page,
        'comp_type': 'RES',
        'nets': {'1': net_a, '2': net_b},
    }

CSA_PAGE_T_WITH_DOT = (
    "FILE_TYPE = MACRO_DRAWING;\n"
    "SET PAGE_NUMBER P1;\n"
    "WIRE 16 -1 (0 0)(100 0);\n"
    "FORCEPROP 2 LAST SIG_NAME T_H\n"
    "WIRE 16 -1 (50 0)(50 100);\n"
    "FORCEPROP 2 LAST SIG_NAME T_V\n"
    "DOT 1 (50 0);\n"
    "CIRCLE 16 -1 (1000 1000)(1100 1000);\n"
)

CSA_PAGE_DOTLESS_CROSS = (
    "FILE_TYPE = MACRO_DRAWING;\n"
    "SET PAGE_NUMBER P2;\n"
    "WIRE 16 -1 (200 0)(300 0);\n"
    "FORCEPROP 2 LAST SIG_NAME CROSS_NO_DOT_H\n"
    "WIRE 16 -1 (250 -50)(250 50);\n"
    "FORCEPROP 2 LAST SIG_NAME CROSS_NO_DOT_V\n"
    "CIRCLE 16 -1 (2000 2000) 150;\n"
)

CSA_PAGE_DOT_CROSS = (
    "FILE_TYPE = MACRO_DRAWING;\n"
    "SET PAGE_NUMBER P3;\n"
    "WIRE 16 -1 (400 0)(500 0);\n"
    "FORCEPROP 2 LAST SIG_NAME CROSS_DOT_H\n"
    "WIRE 16 -1 (450 -50)(450 50);\n"
    "FORCEPROP 2 LAST SIG_NAME CROSS_DOT_V\n"
    "DOT 1 (450 0);\n"
)

CSA_PAGE_SPLIT_CROSS_WITH_ARC = (
    "FILE_TYPE = MACRO_DRAWING;\n"
    "SET PAGE_NUMBER P4;\n"
    "WIRE 16 -1 (600 0)(650 0);\n"
    "FORCEPROP 2 LAST SIG_NAME SPLIT_CROSS_L\n"
    "WIRE 16 -1 (650 0)(700 0);\n"
    "FORCEPROP 2 LAST SIG_NAME SPLIT_CROSS_R\n"
    "WIRE 16 -1 (650 -50)(650 0);\n"
    "FORCEPROP 2 LAST SIG_NAME SPLIT_CROSS_D\n"
    "WIRE 16 -1 (650 0)(650 50);\n"
    "FORCEPROP 2 LAST SIG_NAME SPLIT_CROSS_U\n"
    "DOT 1 (650 0);\n"
    "ARC 16 -1 (3000 3000)(3100 3000)(3050 3050);\n"
)

def sample_part_block():
    return (
        "PART_NAME\n"
        "C1A104 'CAP_HDL-HQ17101005HS0,100NF,10%,0402,X7R,50V':\n"
        "SECTION_NUMBER 1\n"
        " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70"
        "@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1_I17"
        "@HQ_CAP.CAP_HDL(CHIPS)':\n"
        " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70"
        "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page1_i17"
        "@hq_cap.cap_hdl(chips)',\n"
        " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70"
        "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page1_i17"
        "@hq_cap.cap_hdl(chips)',\n"
        " PATH='I17',\n"
        " DRAWING='@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE1',\n"
        " PHYS_PAGE='1',\n"
        " BOM_OPTION='DEPOP',\n"
        " PACKAGE='0402',\n"
        " HQ_CODE='HQ17101005HS0',\n"
        " VALUE='100NF'\n"
    )

def deep_hierarchy_part_block():
    return (
        "PART_NAME\n"
        "C9A001 'CAP_HDL-HQ99999999,10NF,10%,0402,X7R,16V':\n"
        "SECTION_NUMBER 1\n"
        " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE144_I70"
        "@GPU_2SW_BOARD_LIB.I2C_REPEATER_9617_CBB_V3(SCH_1):PAGE3_I17"
        "@GPU_2SW_BOARD_LIB.GRAND_CHILD_BLOCK(SCH_1):PAGE2_I5"
        "@HQ_CAP.CAP_HDL(CHIPS)':\n"
        " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page144_i70"
        "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page3_i17"
        "@gpu_2sw_board_lib.grand_child_block(sch_1):page2_i5"
        "@hq_cap.cap_hdl(chips)',\n"
        " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page114_i70"
        "@gpu_2sw_board_lib.i2c_repeater_9617_cbb_v3(sch_1):page3_i17"
        "@gpu_2sw_board_lib.grand_child_block(sch_1):page2_i5"
        "@hq_cap.cap_hdl(chips)',\n"
        " DRAWING='@GPU_2SW_BOARD_LIB.GRAND_CHILD_BLOCK(SCH_1):PAGE2',\n"
        " HQ_CODE='HQ99999999',\n"
        " VALUE='10NF'\n"
    )

def pex90144_part_block():
    return (
        "PART_NAME\n"
        "C1A101 'CAP_HDL-HQ171010060D0,220NF,10%,0201,X6S,6.3V':\n"
        "REUSE_INSTANCE='PEX90144_CBB_V1A101',\n"
        "REUSE_NAME='PEX90144_CBB_V1',\n"
        "REUSE_PID='906';\n"
        "SECTION_NUMBER 1\n"
        " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE112_I167"
        "@GPU_2SW_BOARD_LIB.PEX90144_CBB_V1(SCH_1):PAGE1_I155"
        "@HQ_CAP.CAP_HDL(CHIPS)':\n"
        " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page112_i167"
        "@gpu_2sw_board_lib.pex90144_cbb_v1(sch_1):page1_i155"
        "@hq_cap.cap_hdl(chips)',\n"
        " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page24_i167"
        "@gpu_2sw_board_lib.pex90144_cbb_v1(sch_1):page1_i155"
        "@hq_cap.cap_hdl(chips)',\n"
        " PATH='I155',\n"
        " DRAWING='@GPU_2SW_BOARD_LIB.PEX90144_CBB_V1(SCH_1):PAGE1',\n"
        " PHYS_PAGE='1',\n"
        " CDS_LIB='hq_cap',\n"
        " CDS_PART_NAME='CAP_HDL-HQ171010060D0,220NF,10%,0201,X6S,6.3V',\n"
        " TOLERANCE='10%',\n"
        " PACKAGE='0201',\n"
        " MATERIAL='X6S',\n"
        " HQ_CODE='HQ171010060D0',\n"
        " VOLTAGE='6.3V',\n"
        " VALUE='220NF',\n"
        " REUSE_PID='906',\n"
        " SUBDESIGN_SUFFIX='101',\n"
        " SUBDESIGN_NAME='PEX90144_CBB_V1',\n"
        " REUSE_INSTANCE='PEX90144_CBB_V1A101',\n"
        " REUSE_NAME='PEX90144_CBB_V1';\n"
    )

def split_symbol_part_block():
    return (
        "PART_NAME\n"
        "U46 'LCMXO3LF_9400C_HDL-HQ11112042009,LCMXO3LF-9400C-5BG484C':;\n"
        "\n"
        "SECTION_NUMBER 1\n"
        " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE152_I2"
        "@HQ_IC.LCMXO3LF_9400C_HDL(CHIPS)':\n"
        " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page152_i2"
        "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
        " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page131_i2"
        "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
        " PATH='I2',\n"
        " DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE152',\n"
        " PHYS_PAGE='131',\n"
        " XY='(-4000,3000)',\n"
        " SPLIT_INST='TRUE',\n"
        " LOCATION='U46',\n"
        " HQ_CODE='HQ11112042009',\n"
        " VALUE='LCMXO3LF-9400C-5BG484C';\n"
        "\n"
        "SECTION_NUMBER 2\n"
        " '@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE151_I98"
        "@HQ_IC.LCMXO3LF_9400C_HDL(CHIPS)':\n"
        " C_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page151_i98"
        "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
        " P_PATH='@gpu_2sw_board_lib.gpu_2sw_board(sch_1):page130_i98"
        "@hq_ic.lcmxo3lf_9400c_hdl(chips)',\n"
        " PATH='I98',\n"
        " DRAWING='@GPU_2SW_BOARD_LIB.GPU_2SW_BOARD(SCH_1):PAGE151',\n"
        " PHYS_PAGE='130',\n"
        " XY='(-4150,2950)',\n"
        " SPLIT_INST='TRUE',\n"
        " LOCATION='U46',\n"
        " HQ_CODE='HQ11112042009',\n"
        " VALUE='LCMXO3LF-9400C-5BG484C';\n"
    )
