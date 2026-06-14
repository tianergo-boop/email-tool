"""
配箱工具核心逻辑模块
处理配箱表.xlsm的读写、配箱计算、ERP/ASN数据导入导出
"""

import os
import json
import re
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string


# ══════════════════════════════════════════════════════
#  常量定义
# ══════════════════════════════════════════════════════

# 配箱表.xlsm 的 sheet 名
SHEET_FORMULA  = "配箱公式"   # ERP订单数据 + 配箱公式
SHEET_BOX      = "配箱表"     # 物流+ASN视图，真正配箱依据
SHEET_SUMMARY  = "汇总"       # 配箱表sheet的透视表
SHEET_ASN      = "ASN"        # ERP导出的ASN数据
SHEET_LOGISTICS = "物流跟踪"   # RPA每日更新的物流数据

# 「配箱公式」sheet 的列（基于实际Excel结构）
# A~Y列：工具自有列（公式、映射、计算结果）
COL_ERP_MATERIAL = "A"   # erp物料
COL_BRAND_MAP    = "H"   # 牌号（公式：VLOOKUP A:B）
COL_BOX_NO       = "J"   # 箱号（核心：VLOOKUP汇总表）
COL_PO           = "K"   # PO
COL_DI           = "S"   # DI（交货指示号）
# Z列开始：ERP原始数据复制区
COL_ORDER_NO     = "Z"   # 客户订单号
COL_STATUS       = "AA"  # 单据状态
COL_LINE_NO      = "AB"  # 行号
COL_SALES_ORG   = "AC"  # 销售组织（如"上海Hub"）
COL_MATERIAL    = "AD"  # 物料名称
COL_DEVICE       = "M"   # 装置（#N/A或装置名）
COL_COA          = "N"   # COA
COL_BATCH        = "BJ"  # 批次号
COL_QTY          = "AI"  # 数量
COL_SHIP_POINT   = "AT"  # 交货指示号（AT列）
COL_INCOTERM     = "X"   # SO Incoterm

# 「汇总」sheet 区域（G~O列，按ETA排序）
SUMMARY_BRAND_COL  = "H"   # 牌号
SUMMARY_HUB_COL   = "B"   # Hub（公式：VLOOKUP DI前缀→Hub名）
SUMMARY_ETA_COL   = "J"   # ETA
SUMMARY_ETD_COL   = "K"   # ETD
SUMMARY_DI_COL     = "M"   # DI
SUMMARY_BOX_COL   = "O"   # 箱号
SUMMARY_SONO_COL  = "D"   # SONO（#N/A=可用）

# 数据起始行（配箱公式sheet，第16行是表头）
DATA_START_ROW = 17


# ══════════════════════════════════════════════════════
#  列映射配置
# ══════════════════════════════════════════════════════

DEFAULT_ERP_MAPPING = {
    "客户订单号": "Z",
    "单据状态": "AA",
    "行号": "AB",
    "销售组织": "AC",
    "物料名称": "AD",
    "LYB陆运承运商": "AE",
    "LYB承运商联系人": "AF",
    "规格型号": "AG",
    "计量单位": "AH",
    "数量": "AI",
    "单价": "AJ",
    "金额": "AK",
    "含税单价": "AL",
    "税额": "AM",
    "价税合计": "AN",
    "发货组织": "AO",
    "客户订单分录序号": "AQ",
    "牌号": "AR",
    "装船要求": "AS",
    "贸易条款": "X",
    "制单日期": "Y",
}


def get_mapping_config_path():
    """获取映射配置文件路径"""
    # 与config.json同目录
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "allocation_mapping.json")


def load_mapping_config():
    """加载列映射配置"""
    path = get_mapping_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"erp_mapping": DEFAULT_ERP_MAPPING}


def save_mapping_config(cfg):
    path = get_mapping_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ══════════════════════════════════════════════════════
#  配箱表读写
# ══════════════════════════════════════════════════════

def load_allocation_workbook(file_path):
    """
    加载配箱表.xlsm
    data_only=True：读取公式的计算结果（不是公式文本）
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")
    return load_workbook(file_path, data_only=True)


def get_box_summary(wb):
    """
    读取「汇总」sheet，返回可用库存池
    汇总表是配箱表sheet的透视表，已筛选SONO=#N/A（可用箱子）
    数据从第6行开始，列映射：
      A=序号, B=Hub, C=运单号, D=重量, E=组内序号,
      H=产品描述(牌号), I=ATD, J=ETA, K=ETD, L=中转港,
      M=DINO(DI), N=运单号, O=箱号, P=重量, Q=批次, R=船公司, S=装置
    """
    ws = wb["汇总"]
    boxes = []

    seen_boxes = set()
    for row in range(6, ws.max_row + 1):
        box_no = ws.cell(row=row, column=15).value  # O列 = 箱号
        if not box_no or str(box_no).strip() == "":
            continue

        box_str = str(box_no).strip()
        if box_str in seen_boxes:
            continue
        seen_boxes.add(box_str)

        hub    = ws.cell(row=row, column=2).value   # B列 = Hub
        brand  = ws.cell(row=row, column=8).value   # H列 = 产品描述/牌号
        di     = ws.cell(row=row, column=13).value  # M列 = DI
        eta    = ws.cell(row=row, column=10).value   # J列 = ETA
        etd    = ws.cell(row=row, column=11).value   # K列 = ETD
        atd    = ws.cell(row=row, column=9).value    # I列 = ATD
        weight = ws.cell(row=row, column=16).value   # P列 = 重量
        batch  = ws.cell(row=row, column=17).value   # Q列 = 批次
        device = ws.cell(row=row, column=19).value   # S列 = 装置
        waybill= ws.cell(row=row, column=3).value    # C列 = 运单号

        # 构造key：数量|销售组织|牌号
        key = f"{weight}|{hub}|{brand}"

        boxes.append({
            "box_no": box_str,
            "brand": str(brand) if brand else "",
            "hub": str(hub) if hub else "",
            "di": str(di) if di else "",
            "eta": eta if isinstance(eta, datetime) else None,
            "etd": etd if isinstance(etd, datetime) else None,
            "atd": atd if isinstance(atd, datetime) else None,
            "weight": weight,
            "batch": str(batch) if batch else "",
            "device": str(device) if device else "",
            "waybill": str(waybill) if waybill else "",
            "key": key,
        })

    # 按ETA排序，附加序号
    boxes.sort(key=lambda x: x["eta"] or datetime(2099, 1, 1))
    for i, b in enumerate(boxes):
        b["seq"] = i + 1

    return boxes


def _find_summary_data_start(ws):
    """找汇总表的数据起始行"""
    for row in range(1, 20):
        # 找包含"箱号"或第一个非空箱号的行
        for col in range(1, 20):
            val = ws.cell(row=row, column=col).value
            if val and "箱号" in str(val):
                return row + 1
    return 6  # 默认第6行


def get_current_orders(wb):
    """
    读取「配箱公式」sheet，返回当前已导入的订单及配箱结果
    返回：[{
        "row": 17,
        "order_no": "4900941331",
        "sales_org": "上海Hub",
        "material": "...",
        "brand": "...",
        "qty": 27,
        "box_no": "CSNU6361037" or None,  # J列结果
        "di": "...",
        "batch": "...",
        "device": "...",
        "status": "已配"|"待配"|"指定"|"无库存",
    }, ...]
    """
    ws = wb["配箱公式"]
    orders = []

    for row in range(2, ws.max_row + 1):
        order_no = ws.cell(row=row, column=26).value  # Z列 = 客户订单号
        if not order_no:
            continue

        order_str = str(int(order_no)) if isinstance(order_no, (int, float)) else str(order_no).strip()

        sales_org = ws.cell(row=row, column=29).value  # AC列 = 销售组织
        material = ws.cell(row=row, column=30).value  # AD列 = 物料名称
        brand     = ws.cell(row=row, column=8).value    # H列 = 牌号
        qty       = ws.cell(row=row, column=35).value  # AI列 = 数量
        box_no    = ws.cell(row=row, column=10).value   # J列 = 箱号
        di        = ws.cell(row=row, column=19).value    # S列 = DI
        device    = ws.cell(row=row, column=13).value   # M列 = 装置
        batch     = ws.cell(row=row, column=62).value   # BJ列 = 批次号
        ship_point= ws.cell(row=row, column=46).value   # AT列 = 交货指示号

        # 判断状态
        if box_no and str(box_no).strip() and str(box_no).strip() != "#N/A" and str(box_no).strip() != "/":
            # 检查J列是公式结果还是手动指定（通过读取带公式的版本判断）
            status = "已配"
        elif brand and str(brand).strip() != "#N/A":
            status = "待配"
        else:
            status = "无库存"

        orders.append({
            "row": row,
            "order_no": order_str,
            "sales_org": str(sales_org) if sales_org else "",
            "material": str(material) if material else "",
            "brand": str(brand) if brand else "",
            "qty": qty,
            "box_no": str(box_no).strip() if box_no and str(box_no).strip() not in ("#N/A", "/") else None,
            "di": str(di).strip() if di and str(di).strip() != "#N/A" else "",
            "batch": str(batch).strip() if batch and str(batch).strip() != "#N/A" else "",
            "device": str(device).strip() if device and str(device).strip() != "#N/A" else "",
            "ship_point": str(ship_point).strip() if ship_point and str(ship_point).strip() != "#N/A" else "",
            "status": status,
        })

    return orders


# ══════════════════════════════════════════════════════
#  配箱计算逻辑（工具自主重算）
# ══════════════════════════════════════════════════════

def allocate_boxes(orders, boxes):
    """
    配箱计算：按行顺序，为每个订单分配可用箱子
    orders: get_current_orders 返回的列表（会被修改）
    boxes: get_box_summary 返回的可用库存池（会被修改）
    返回：(allocated_orders, remaining_boxes)
    """
    # 构建可用库存池（按key分组，组内按ETA排序）
    pool = {}  # key -> [box, ...] 按ETA排序
    for b in boxes:
        k = b["key"]
        if k not in pool:
            pool[k] = []
        pool[k].append(b)

    allocated = []
    used_boxes = set()

    for order in orders:
        if order["status"] == "指定" and order.get("_manual_box"):
            # 手动指定的，跳过自动分配
            allocated.append(order)
            used_boxes.add(order["_manual_box"])
            continue

        key = f"{order['qty']}|{order['sales_org']}|{order['brand']}"
        if key not in pool or not pool[key]:
            order["status"] = "无库存"
            allocated.append(order)
            continue

        # 取ETA最早的可用的箱子
        box = pool[key].pop(0)  # 从池子里取出
        order["box_no"] = box["box_no"]
        order["status"] = "已配"
        order["_allocated_from"] = box
        used_boxes.add(box["box_no"])
        allocated.append(order)

    remaining = [b for b in boxes if b["box_no"] not in used_boxes]
    return allocated, remaining


def manual_assign_box(orders, boxes, order_rows, box_no):
    """
    手动指定箱号
    order_rows: 要指定的订单行号列表
    box_no: 指定的箱号
    返回：(updated_orders, error_msg)
    """
    # 检查箱子是否在可用池里
    box = None
    for b in boxes:
        if b["box_no"] == box_no:
            box = b
            break

    if not box:
        return None, f"箱号 {box_no} 不在可用库存池中"

    # 为选中的订单分配
    for order in orders:
        if order["row"] in order_rows:
            order["box_no"] = box_no
            order["status"] = "指定"
            order["_manual_box"] = box_no

    return orders, None


def manual_assign_di(orders, boxes, order_rows, di):
    """
    手动指定DI：该DI下的箱子按ETA先到先配
    """
    # 找出该DI下的所有箱子，按ETA排序
    di_boxes = [b for b in boxes if b["di"] == di]
    di_boxes.sort(key=lambda x: x["eta"] or datetime(2099, 1, 1))

    if not di_boxes:
        return None, f"DI {di} 下没有可用的箱子"

    # 按序分配
    order_idx = 0
    for order in orders:
        if order["row"] in order_rows:
            if order_idx < len(di_boxes):
                box = di_boxes[order_idx]
                order["box_no"] = box["box_no"]
                order["status"] = "指定"
                order["_manual_box"] = box["box_no"]
                order_idx += 1
            else:
                order["status"] = "无库存"

    return orders, None


# ══════════════════════════════════════════════════════
#  ERP订单导入
# ══════════════════════════════════════════════════════

def import_erp_orders(erp_file_path, allocation_file_path, mapping=None):
    """
    导入ERP订单到配箱表.xlsm
    erp_file_path: ERP导出的Excel
    allocation_file_path: 配箱表.xlsm路径
    mapping: 列映射配置（None则使用默认）
    返回：(success_count, error_msg)
    """
    if mapping is None:
        cfg = load_mapping_config()
        mapping = cfg.get("erp_mapping", DEFAULT_ERP_MAPPING)

    # 读取ERP文件
    erp_wb = load_workbook(erp_file_path, data_only=True)
    erp_ws = erp_wb.active

    # 找ERP表头行（第一行包含"客户订单号"或"订单"）
    header_row = _find_erp_header(erp_ws)
    if header_row == 0:
        return 0, "无法识别ERP文件表头，请检查文件格式"

    # 构建ERP列名→列索引的映射
    erp_col_map = {}
    for col in range(1, erp_ws.max_column + 1):
        val = erp_ws.cell(row=header_row, column=col).value
        if val:
            erp_col_map[str(val).strip()] = col

    # 打开配箱表（带公式版本，用于写入）
    alloc_wb = load_workbook(allocation_file_path)
    alloc_ws = alloc_wb["配箱公式"]

    # 找配箱表的下一个空行
    next_row = _find_next_data_row(alloc_ws)

    # 写入数据
    success_count = 0
    for data_row in range(header_row + 1, erp_ws.max_row + 1):
        # 检查是否有客户订单号（必须有才写入）
        order_no_col = erp_col_map.get("客户订单号")
        if not order_no_col:
            break
        order_val = erp_ws.cell(row=data_row, column=order_no_col).value
        if not order_val:
            continue

        # 按映射写入配箱公式sheet的Z列及以后
        for erp_col_name, alloc_col in mapping.items():
            erp_col_idx = erp_col_map.get(erp_col_name)
            if not erp_col_idx:
                continue
            val = erp_ws.cell(row=data_row, column=erp_col_idx).value
            if val is None:
                continue

            alloc_col_idx = column_index_from_string(alloc_col)
            alloc_ws.cell(row=next_row, column=alloc_col_idx).value = val

        # 同时写入A列（erp物料=物料名称）
        material = erp_ws.cell(row=data_row, column=erp_col_map.get("物料名称", 0)).value
        if material:
            alloc_ws.cell(row=next_row, column=1).value = material  # A列

        next_row += 1
        success_count += 1

    # 保存
    alloc_wb.save(allocation_file_path)
    erp_wb.close()

    return success_count, None


def _find_erp_header(ws):
    """找ERP文件的表头行"""
    for row in range(1, 10):
        for col in range(1, 20):
            val = ws.cell(row=row, column=col).value
            if val and ("客户订单号" in str(val) or "订单号" in str(val)):
                return row
    return 0


def _find_next_data_row(ws):
    """找配箱公式sheet的下一个空行"""
    for row in range(DATA_START_ROW, ws.max_row + 2):
        val = ws.cell(row=row, column=26).value  # Z列 = 客户订单号
        if not val:
            return row
    return ws.max_row + 1


# ══════════════════════════════════════════════════════
#  ASN导入 + 运单号补录
# ══════════════════════════════════════════════════════

def import_asn(asn_file_path, allocation_file_path):
    """
    导入ASN数据到配箱表.xlsm
    """
    asn_wb = load_workbook(asn_file_path, data_only=True)
    asn_ws = asn_wb.active

    alloc_wb = load_workbook(allocation_file_path)
    alloc_ws = alloc_wb["ASN"]

    # 清空ASN sheet（保留表头）
    for row in range(2, alloc_ws.max_row + 1):
        for col in range(1, alloc_ws.max_column + 1):
            alloc_ws.cell(row=row, column=col).value = None

    # 复制ASN数据
    for row in range(1, asn_ws.max_row + 1):
        for col in range(1, asn_ws.max_column + 1):
            val = asn_ws.cell(row=row, column=col).value
            alloc_ws.cell(row=row, column=col).value = val

    alloc_wb.save(allocation_file_path)
    asn_wb.close()
    return True, None


def supplement_waybill(allocation_file_path, box_no, waybill_no):
    """
    补录运单号：在ASN sheet中找对应箱号，填入运单号
    """
    wb = load_workbook(allocation_file_path)
    ws = wb["ASN"]

    found = False
    for row in range(1, ws.max_row + 1):
        bn = ws.cell(row=row, column=1).value  # 假设第1列是箱号
        if bn and str(bn).strip() == box_no:
            # 找运单号列（假设是某列）
            # 需根据实际ASN格式确认
            ws.cell(row=row, column=2).value = waybill_no
            found = True
            break

    wb.save(allocation_file_path)
    return found, None


# ══════════════════════════════════════════════════════
#  导出
# ══════════════════════════════════════════════════════

def export_allocation_list(orders, output_path):
    """
    导出配箱清单（箱号+PO）用于ERP导入
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "配箱清单"

    # 表头
    ws["A1"] = "箱号"
    ws["B1"] = "PO"
    ws["C1"] = "客户订单号"
    ws["D1"] = "销售组织"
    ws["E1"] = "牌号"
    ws["F1"] = "数量"

    row = 2
    for order in orders:
        if order["box_no"]:
            ws.cell(row=row, column=1).value = order["box_no"]
            ws.cell(row=row, column=2).value = order.get("po", "")
            ws.cell(row=row, column=3).value = order["order_no"]
            ws.cell(row=row, column=4).value = order["sales_org"]
            ws.cell(row=row, column=5).value = order["brand"]
            ws.cell(row=row, column=6).value = order["qty"]
            row += 1

    wb.save(output_path)
    return output_path


# ══════════════════════════════════════════════════════
#  工具函数
# ══════════════════════════════════════════════════════

def format_eta(dt):
    """格式化ETA/ETD为字符串"""
    if isinstance(dt, datetime):
        return dt.strftime("%Y/%m/%d")
    return str(dt) if dt else ""


def get_hub_from_di(di_str):
    """从DI前缀获取Hub（如"SH"→"上海Hub"）"""
    if not di_str or len(di_str) < 2:
        return ""
    prefix = di_str[:2].upper()
    hub_map = {
        "SH": "上海Hub",
        "HP": "黄埔Hub",
        "QD": "青岛Hub",
        "NB": "宁波Hub",
        "ST": "汕头Hub",
        "XG": "新港Hub",
        "XM": "厦门Hub",
        "NS": "南沙Hub",
        "WF": "潍坊Hub",
        "NJ": "南京Hub",
        "TC": "太仓Hub",
        "SQ": "宿迁Hub",
        "CQ": "重庆Hub",
        "WH": "武汉Hub",
        "TZ": "台州Hub",
        "HF": "合肥Hub",
        "FZ": "福州Hub",
    }
    return hub_map.get(prefix, "")
