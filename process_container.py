'''
ECOUNT ERP 货柜费用分摊核算与全套导入表生成引擎
包含：
1. 一般凭证上传格式 (13列)
2. 采购入库格式 (25列)
3. 新品目格式 (17列)
4. 品目编码主数据 (11列)
5. 货柜三费分摊明细表 (14列)
'''
import sys
import os
import re
import json
import copy
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    
    try:
        sys.stdout.reconfigure(encoding = 'utf-8', errors = 'replace')
    except Exception:
        pass

if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    
    try:
        sys.stderr.reconfigure(encoding = 'utf-8', errors = 'replace')
    except Exception:
        pass


def safe_print(msg):
    '''安全打印函数，防止 Windows GBK 编码报错'''
    
    try:
        print(msg)
        return None
    except Exception:
        
        try:
            enc = sys.stdout.encoding or 'gbk'
            print(str(msg).encode(enc, errors = 'replace').decode(enc))
        except Exception:
            pass

        return None



def sanitize_item_code(name_or_model):
    """
    品目编码清洗函数：
    ECOUNT ERP 规定编码中不允许包含空格，如果原数据名称/型号中含有空格，
    自动使用连字符 '-' 代替空格作为品目编码。
    例如: 'MMA-300S IGBT' -> 'MMA-300S-IGBT', 'SAMPLE ITEM' -> 'SAMPLE-ITEM'
    """
    if not name_or_model:
        return ''
    s = str(name_or_model).strip()
    code = re.sub('\\s+', '-', s)
    return code

DEFAULT_COUNTERPARTY_MAP = {
    '优势库': {
        'rmb_code': '31750002',
        'rmb_name': '浙江优势库进出口有限公司',
        'usd_code': '31750003',
        'usd_name': '浙江优势库进出口有限公司USD' },
    '颐顿': {
        'rmb_code': '33108107',
        'rmb_name': '浙江颐顿机电有限公司',
        'usd_code': '3310810701',
        'usd_name': '浙江颐顿机电有限公司usd' } }
DEFAULT_TEMPLATE_CONFIG = {
    'purchase': 'C:\\Users\\123\\Downloads\\采购入库格式.xlsx',
    'new_item': 'C:\\Users\\123\\Downloads\\新增品目格式.xlsx',
    'voucher': 'C:\\Users\\123\\Downloads\\一般凭证格式.xlsx',
    'export_individual_files': False,
    'export_bundle_workbook': True }
PURCHASE_TEMPLATE_FIELD_DEFS = {
    'date': {
        'label': '日期',
        'keywords': [
            '日期',
            '单据日期',
            '交易日期',
            'date'] },
    'seq_no': {
        'label': '序号',
        'keywords': [
            '序号',
            '序列',
            'seq',
            'no'] },
    'co_code': {
        'label': '往来单位编码',
        'keywords': [
            '往来单位编码',
            '往来单位编号',
            '客户编码',
            '供应商编码'] },
    'co_name': {
        'label': '往来单位名',
        'keywords': [
            '往来单位名',
            '往来单位名称',
            '往来单位',
            '客户名称'] },
    'staff': {
        'label': '职员',
        'keywords': [
            '职员',
            '承办人',
            '业务员',
            '担当者'] },
    'warehouse': {
        'label': '接收仓库',
        'keywords': [
            '接收仓库',
            '仓库/工厂编码',
            '进货仓库',
            '仓库编码',
            '仓库'] },
    'trade_type': {
        'label': '交易类型',
        'keywords': [
            '交易类型',
            '税率类型',
            '类型'] },
    'currency': {
        'label': '货币',
        'keywords': [
            '货币',
            '外币编码',
            '币种'] },
    'rate': {
        'label': '汇率',
        'keywords': [
            '汇率',
            '外币汇率'] },
    'contract_no': {
        'label': '合同号',
        'keywords': [
            '合同号',
            '合同编号',
            '合同',
            '文本类型2'] },
    'container_no': {
        'label': '集装箱号码',
        'keywords': [
            '集装箱号码',
            '集装箱号',
            '柜号',
            '箱号',
            '文本类型1'] },
    'rate_text': {
        'label': '汇率(文字)',
        'keywords': [
            '汇率',
            '文字型',
            '附加信息'] },
    'ship_date': {
        'label': '发货日期',
        'keywords': [
            '发货日期',
            '出货日期',
            '日期型YYYYMMDD',
            '出运日期'] },
    'item_code': {
        'label': '品目编码',
        'keywords': [
            '品目编码',
            '品目编号',
            '商品编码',
            '物料编码'] },
    'item_name': {
        'label': '品目名',
        'keywords': [
            '品目名',
            '品目名称',
            '商品名称',
            '物料名称'] },
    'spec': {
        'label': '规格名',
        'keywords': [
            '规格名',
            '规格',
            '型号规格',
            '规格型号'] },
    'qty': {
        'label': '数量',
        'keywords': [
            '数量',
            '入库数量',
            '出货数量'] },
    'tax_price': {
        'label': '序列号/批号',
        'keywords': [
            '序列号/批号',
            '批号',
            '序列号',
            '批次'] },
    'untaxed_price': {
        'label': '摘要',
        'keywords': [
            '摘要',
            '备注',
            '说明'] },
    'usd_total': {
        'label': '增值税',
        'keywords': [
            '增值税',
            '税额',
            '进项税'] },
    'untaxed_total': {
        'label': '附带费用',
        'keywords': [
            '附带费用',
            '附加费用',
            '分摊费用',
            '三费分摊',
            '四费分摊'] },
    'fee_alloc': {
        'label': '供应价',
        'keywords': [
            '供应价',
            '税前金额',
            '未税总价',
            '未税金额'] },
    'tax': {
        'label': '外币金额',
        'keywords': [
            '外币金额',
            '美金金额',
            '外币'] },
    'remark': {
        'label': '单价',
        'keywords': [
            '单价',
            '不含税单价',
            '未税单价'] },
    'batch_no': {
        'label': '单价(包括税)',
        'keywords': [
            '单价(包括税)',
            '单价(含税)',
            '含税单价'] } }
NEW_ITEM_TEMPLATE_FIELD_DEFS = {
    'item_code': {
        'label': '品目编码',
        'keywords': [
            '品目编码',
            '品目编号',
            '商品编码'] },
    'item_name': {
        'label': '品目名',
        'keywords': [
            '品目名',
            '品目名称',
            '商品名称'] },
    'spec_type': {
        'label': '规格类型',
        'keywords': [
            '规格类型'] },
    'spec': {
        'label': '规格',
        'keywords': [
            '规格',
            '规格名',
            '规格信息'] },
    'unit': {
        'label': '单位',
        'keywords': [
            '单位',
            '计量单位'] },
    'item_type': {
        'label': '品目类型',
        'keywords': [
            '品目类型',
            '品目属性'] },
    'is_kit': {
        'label': '使用套件',
        'keywords': [
            '使用套件',
            '是否套件',
            '套件'] },
    'qty_mgmt': {
        'label': '数量管理',
        'keywords': [
            '数量管理',
            '是否数量管理'] },
    'in_price': {
        'label': '入库单价',
        'keywords': [
            '入库单价',
            '采购单价'] },
    'in_price_tax': {
        'label': '入库单价 是否包括增值税',
        'keywords': [
            '入库单价 是否包括增值税',
            '入库单价是否含税',
            '入库含税'] },
    'process': {
        'label': '生产流程',
        'keywords': [
            '生产流程',
            '工序流程',
            '流程'] },
    'out_price': {
        'label': '出库单价',
        'keywords': [
            '出库单价',
            '销售单价'] },
    'out_price_tax': {
        'label': '出库单价 是否包括增值税',
        'keywords': [
            '出库单价 是否包括增值税',
            '出库单价是否含税',
            '出库含税'] },
    'comb1': {
        'label': '品目组合1',
        'keywords': [
            '品目组合1',
            '组合1',
            '品目分类1'] },
    'pack_qty': {
        'label': '装数',
        'keywords': [
            '装数',
            '数字型附加事项1',
            '附加事项1',
            '每箱装数'] },
    'unit_cbm': {
        'label': '体积',
        'keywords': [
            '体积',
            '数字型附加事项2',
            '附加事项2',
            '单台体积'] },
    'unit_weight': {
        'label': '重量',
        'keywords': [
            '重量',
            '数字型附加事项3',
            '附加事项3',
            '单台重量'] } }
VOUCHER_TEMPLATE_FIELD_DEFS = {
    'date': {
        'label': '凭证日期',
        'keywords': [
            '凭证日期',
            '单据日期',
            '日期',
            'date'] },
    'seq_no': {
        'label': '序号',
        'keywords': [
            '序号',
            '序列',
            'seq',
            'no'] },
    'voucher_no': {
        'label': '会计凭证No.',
        'keywords': [
            '会计凭证No.',
            '会计凭证',
            '会计No.',
            '凭证号'] },
    'dr_cr': {
        'label': '类型',
        'keywords': [
            '类型',
            '借贷类型',
            '借/贷',
            '借贷'] },
    'account_code': {
        'label': '科目编码',
        'keywords': [
            '科目编码',
            '科目编号',
            '会计科目',
            '科目'] },
    'co_code': {
        'label': '往来单位编码',
        'keywords': [
            '往来单位编码',
            '往来单位编号',
            '客户编码',
            '单位编码'] },
    'co_name': {
        'label': '往来单位名',
        'keywords': [
            '往来单位名',
            '往来单位名称',
            '往来单位',
            '客户名称'] },
    'amount': {
        'label': '金额',
        'keywords': [
            '金额',
            '借贷金额',
            '本币金额',
            '人民币金额'] },
    'usd_amount': {
        'label': '外币金额',
        'keywords': [
            '外币金额',
            '美金金额',
            '外币'] },
    'rate': {
        'label': '汇率',
        'keywords': [
            '汇率',
            '美金汇率'] },
    'summary_code': {
        'label': '摘要编码',
        'keywords': [
            '摘要编码',
            '摘要编号'] },
    'summary': {
        'label': '摘要',
        'keywords': [
            '摘要',
            '凭证摘要',
            '业务摘要'] },
    'dept': {
        'label': '部门',
        'keywords': [
            '部门',
            '部门编码',
            '部门名称'] },
    'project': {
        'label': '项目记事板',
        'keywords': [
            '项目记事板',
            '项目',
            '工程',
            '项目编码'] } }

def parse_comment_details(comment_text):
    '''
    智能解析 ECOUNT ERP 表头批注文本
    提取：业务说明、字数/位数限制、枚举值代码说明
    '''
    if not comment_text:
        return {
            'desc': '',
            'limit_desc': '',
            'enum_desc': '',
            'raw': '' }
    raw = str(comment_text).strip()
    lines = [ l.strip() for l in raw.splitlines() if l.strip() ]
    desc = lines[0] if lines else ''
    limit_parts = []
    enum_parts = []
    for l in lines:
        clean_l = l.replace('*', '').replace('-', '').strip()
        if not (not ('最多' in clean_l) and not ('字数限制' in clean_l) and not ('文字限制' in clean_l) and not ('YYYYMMDD' in clean_l) and not ('位' in clean_l)) or '字' in clean_l:
            if not ('(' in clean_l and ')' in clean_l) or not ('/' in clean_l):
                if 'YYYYMMDD' in clean_l:
                    limit_parts.append('YYYYMMDD (8位)')
                elif clean_l and clean_l not in ('输入字数限制', '输入文字限制', '字数限制', 'Character Limit'):
                    limit_parts.append(clean_l)
        if '(' in clean_l and ')' in clean_l:
            if not (not ('/' in clean_l) and not ('0' in clean_l) and not ('1' in clean_l) and not ('2' in clean_l) and not ('3' in clean_l)) or '4' in clean_l:
                enum_parts.append(clean_l)
    if enum_parts:
        return {
            'desc': desc,
            'limit_desc': ' | '.join(limit_parts) if limit_parts else '无特殊限制',
            'enum_desc': ' | '.join(enum_parts),
            'raw': raw }
    return {
        'desc': desc,
        'limit_desc': ' | '.join(limit_parts) if limit_parts else '无特殊限制',
        'enum_desc': '',
        'raw': raw }


def inspect_template_file(file_path, template_type = None):
    '''
    深度解析 Excel 上传格式模板文件：
    提取 Sheet 名称、所有表头、每个单元格的批注内容与解析规则
    '''
    if not file_path or not os.path.exists(file_path):
        return {
            'exists': False,
            'file_path': file_path,
            'error': '文件不存在',
            'sheet_name': '',
            'max_columns': 0,
            'columns': [] }
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only = True)
        ws = None
        for sname in wb.sheetnames:
            if template_type == 'purchase':
                if '采购' in sname or '入库' in sname:
                    ws = wb[sname]
                    break
            if template_type == 'new_item':
                if '品目' in sname or '新品' in sname:
                    ws = wb[sname]
                    break
            if template_type == 'voucher' and '凭证' in sname:
                ws = wb[sname]
                break
        if ws is None:
            ws = wb.active
        cols_info = []
        for c_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row = 1, column = c_idx)
            header_val = str(cell.value or '').strip()
            cmt_text = cell.comment.text if cell.comment else ''
            rule_info = parse_comment_details(cmt_text)
            cols_info.append({
                'col_idx': c_idx,
                'coord': cell.coordinate,
                'header': header_val,
                'has_comment': bool(cell.comment),
                'comment': cmt_text,
                'desc': rule_info['desc'],
                'limit_desc': rule_info['limit_desc'],
                'enum_desc': rule_info['enum_desc'] })
        sheet_name = ws.title
        max_cols = ws.max_column
        wb.close()
        return {
            'exists': True,
            'file_path': file_path,
            'error': None,
            'sheet_name': sheet_name,
            'max_columns': max_cols,
            'columns': cols_info }
    except Exception as e:
        return {
            'exists': True,
            'file_path': file_path,
            'error': str(e),
            'sheet_name': '',
            'max_columns': 0,
            'columns': [] }



def load_template_config():
    """从本地 JSON 文件读取上传格式模板配置"""
    cfg_file = get_config_path('template_config.json')
    if os.path.exists(cfg_file):
        try:
            with open(cfg_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    merged = dict(DEFAULT_TEMPLATE_CONFIG)
                    merged.update(data)
                    return merged
        except Exception as e:
            safe_print(f'[Warn] 读取模板配置失败: {e}')
    cfg = dict(DEFAULT_TEMPLATE_CONFIG)
    user_dl = os.path.expanduser('~\\Downloads')
    cand_purchase = [
        os.path.join(user_dl, '采购入库格式.xlsx'),
        os.path.join(user_dl, 'Template.xlsx'),
        os.path.join(get_base_dir(), '采购入库格式.xlsx')
    ]
    cand_new_item = [
        os.path.join(user_dl, '新增品目格式.xlsx'),
        os.path.join(user_dl, 'Template (2).xlsx'),
        os.path.join(user_dl, 'Template (1).xlsx'),
        os.path.join(get_base_dir(), '新增品目格式.xlsx')
    ]
    cand_voucher = [
        os.path.join(user_dl, '一般凭证格式.xlsx'),
        os.path.join(user_dl, 'Template_一般凭证.xlsx'),
        os.path.join(get_base_dir(), '一般凭证格式.xlsx')
    ]
    for p in cand_purchase:
        if os.path.exists(p):
            cfg['purchase'] = p
            break
    for p in cand_new_item:
        if os.path.exists(p):
            cfg['new_item'] = p
            break
    for p in cand_voucher:
        if os.path.exists(p):
            cfg['voucher'] = p
            break
    return cfg


def save_template_config(cfg):
    """保存上传格式模板配置到本地 JSON"""
    cfg_file = get_config_path('template_config.json')
    try:
        with open(cfg_file, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        safe_print(f'[Error] 保存模板配置失败: {e}')
        return False


DEFAULT_EXISTING_MASTER = [
    ('A0001', 'SAMPLE ITEM', '[商品]', '', '', '', 'YES', 1, '', '', ''),
    ('JQB001', '机器设备', '[商品]', '50*60*60', '', '', 'YES', 1, '', '', ''),
    ('RWD-12000E', 'RWD-12000E', '[商品]', '710*560*630', '', '', 'YES', 1, 1, 0.2505, 94),
    ('RWD-15000E', 'RWD-15000E', '[商品]', '710*560*630', '', '', 'YES', 1, 1, 0.2505, 98.2),
    ('RWD-5200E', 'RWD-5200E', '[商品]', '610*450*470', '', '', 'YES', 1, 1, 0.129, 49.4),
    ('RWD-6500E', 'RWD-6500E', '[商品]', '710*560*630', '', '', 'YES', 1, 1, 0.2505, 87),
    ('RWD-7500E', 'RWD-7500E', '[商品]', '710*560*630', '', '', 'YES', 1, 1, 0.2505, 89.2),
    ('RWD-8800E', 'RWD-8800E', '[商品]', '710*560*630', '', '', 'YES', 1, 1, 0.2505, 93.4),
    ('MINI250S', 'MINI250S', '[商品]', 'MINI250S', '', '', 'YES', 1, 6, 0.009559, ''),
    ('MINI300S', 'MINI300S', '[商品]', 'MINI300S', '', '', 'YES', 1, 6, 0.009853, ''),
    ('TB-250', 'TB-250', '[商品]', 'TB-250', '', '', 'YES', 1, 6, 0.008873, ''),
    ('TB-300', 'TB-300', '[商品]', 'TB-300', '', '', 'YES', 1, 4, 0.0147, ''),
    ('TB-400', 'TB-400', '[商品]', 'TB-400', '', '', 'YES', 1, 4, 0.0147, ''),
    ('ARC-400S', 'ARC-400S', '[商品]', 'ARC-400S', '', '', 'YES', 1, 2, 0.027979, ''),
    ('ARC-500S', 'ARC-500S', '[商品]', 'ARC-500S', '', '', 'YES', 1, 2, 0.028, ''),
    ('ARC-630S', 'ARC-630S', '[商品]', 'ARC-630S', '', '', 'YES', 1, 2, 0.0288, ''),
    ('MMA-300S-IGBT', 'MMA-300S IGBT', '[商品]', 'MMA-300S IGBT', '', '', 'YES', 1, 2, 0.031367, ''),
    ('MMA-300S-MOS', 'MMA-300S MOS', '[商品]', 'MMA-300S MOS', '', '', 'YES', 1, 2, 0.0288, ''),
    ('ARC-4000', 'ARC-4000', '[商品]', 'ARC-4000', '', '', 'YES', 1, 1, 0.149, ''),
    ('ARC-5000', 'ARC-5000', '[商品]', 'ARC-5000', '', '', 'YES', 1, 1, 0.149, ''),
    ('CR-9525A', 'CR-9525A', '[商品]', 'CR-9525A', '', '', 'YES', 1, 4, 0.0234, ''),
    ('CR-9550A', 'CR-9550A', '[商品]', 'CR-9550A', '', '', 'YES', 1, 4, 0.0272, '')]
SUMMARY_FIELD_DEFS = {
    'seq_no': {
        'label': '序号/编号',
        'keywords': [
            '序号',
            '编号',
            'no',
            'id'] },
    'trade_co': {
        'label': '外贸公司',
        'keywords': [
            '外贸公司',
            '贸易公司',
            '往来单位',
            '出运公司',
            '抬头',
            '公司',
            '单位'] },
    'contract_no': {
        'label': '合同号',
        'keywords': [
            '合同号',
            '合同',
            'contract',
            'contractno'] },
    'date': {
        'label': '出货日期',
        'keywords': [
            '出货日期',
            '发货日期',
            '出运日期',
            '日期',
            'date'] },
    'bl_no': {
        'label': '提单号',
        'keywords': [
            '提单号',
            '提单',
            'bl',
            'b/l'] },
    'container_no': {
        'label': '集装箱号码',
        'keywords': [
            '集装箱号码',
            '集装箱号',
            '集装箱',
            '柜号',
            '箱号',
            'container'] },
    'boxes': {
        'label': '总箱数',
        'keywords': [
            '出货数量（箱）',
            '出货数量(箱)',
            '总箱数',
            '箱数',
            '总件数',
            'boxes'] },
    'total_price': {
        'label': '总价(RMB)',
        'keywords': [
            '总价',
            '人民币总价',
            '金额',
            '总金额'] },
    'usd_total': {
        'label': '报关美金(USD)',
        'keywords': [
            '报关美金',
            '美金总价',
            '美金金额',
            'usd'] },
    'ocean_usd': {
        'label': '海运费(USD)',
        'keywords': [
            '海运费',
            '海运',
            'ocean',
            'freight'] },
    'truck_rmb': {
        'label': '内陆托箱费(RMB)',
        'keywords': [
            '内陆托箱费',
            '托箱费',
            '内托费',
            '内陆拖车',
            '拖车费',
            '内陆运费',
            'truck'] },
    'insurance_usd': {
        'label': '保费(USD)',
        'keywords': [
            '保费（美金）',
            '保费(美金)',
            '保费',
            '保险费',
            '保险',
            'insurance'] },
    'agency_rmb': {
        'label': '代理费/3个点(RMB)',
        'keywords': [
            '3个点的费用',
            '三个点',
            '代理费',
            '点费',
            '服务费',
            'agency'] },
    'cbm': {
        'label': '总体积(CBM)',
        'keywords': [
            '总体积',
            '体积',
            'cbm'] },
    'rate': {
        'label': '汇率',
        'keywords': [
            '汇率',
            '美金汇率',
            'rate'] } }
DETAIL_FIELD_DEFS = {
    'seq_no': {
        'label': '编号/序号',
        'keywords': [
            '编号',
            '序号',
            'no'] },
    'contract_no': {
        'label': '合同号',
        'keywords': [
            '合同号',
            '合同',
            'contract'] },
    'date': {
        'label': '出货日期',
        'keywords': [
            '出货日期',
            '发货日期',
            '日期',
            'date'] },
    'bl_no': {
        'label': '提单号',
        'keywords': [
            '提单号',
            '提单',
            'bl',
            'b/l'] },
    'container_no': {
        'label': '集装箱号码',
        'keywords': [
            '集装箱号码',
            '集装箱号',
            '集装箱',
            '柜号',
            '箱号',
            'container'] },
    'trade_co': {
        'label': '外贸公司',
        'keywords': [
            '外贸公司',
            '贸易公司',
            '往来单位',
            '出运公司',
            '公司',
            '单位',
            '抬头',
            '供货商'] },
    'cat': {
        'label': '品名/分类',
        'keywords': [
            '品名',
            '商品名称',
            '货物名称',
            '品类',
            '类别'] },
    'model': {
        'label': '型号/品目',
        'keywords': [
            '型号',
            '规格型号',
            '品目编码',
            '品目名',
            '品目',
            '物料型号',
            'model'] },
    'boxes': {
        'label': '箱数',
        'keywords': [
            '箱数',
            '件数',
            '箱'] },
    'pack_qty': {
        'label': '装数(台/箱)',
        'keywords': [
            '装数',
            '每箱装数',
            '每箱台数',
            '包装数',
            'pack'] },
    'qty': {
        'label': '出货数量(台)',
        'keywords': [
            '出货数量',
            '数量',
            '出货数量(台)',
            '台数',
            '数量(台)',
            'qty'] },
    'unit_price': {
        'label': '含税单价(RMB)',
        'keywords': [
            '单价',
            '含税单价',
            '采购单价',
            '人民币单价',
            '价格',
            'price'] },
    'total_price': {
        'label': '含税总价(RMB)',
        'keywords': [
            '总价',
            '含税总价',
            '采购总额',
            '人民币总价',
            'total'] },
    'usd_total': {
        'label': '美金总价(USD)',
        'keywords': [
            '美金总价',
            '美金金额',
            '报关美金',
            'usd'] },
    'untaxed_unit': {
        'label': '未税单价(RMB)',
        'keywords': [
            '甲骨文不含税单价',
            '不含税单价',
            '未税单价'] },
    'untaxed_total': {
        'label': '未税总价(RMB)',
        'keywords': [
            '甲骨文不含税总价',
            '不含税总价',
            '未税总价',
            '未税金额'] },
    'agency_rmb': {
        'label': '代理费/3个点(RMB)',
        'keywords': [
            '3个点的费用',
            '三个点',
            '代理费'] },
    'cbm': {
        'label': '体积(CBM)',
        'keywords': [
            '体积',
            'cbm'] } }

def get_base_dir():
    '''获取程序运行根目录（兼容 PyInstaller 打包环境）'''
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_config_path(filename):
    return os.path.join(get_base_dir(), filename)


def load_counterparty_map():
    """从本地 JSON 文件读取往来单位映射配置"""
    cfg_file = get_config_path('counterparty_config.json')
    if os.path.exists(cfg_file):
        try:
            with open(cfg_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            safe_print(f'[Warn] 读取往来单位配置失败: {e}')
            return dict(DEFAULT_COUNTERPARTY_MAP)
    return dict(DEFAULT_COUNTERPARTY_MAP)


def save_counterparty_map(mapping):
    """保存往来单位配置到本地 JSON"""
    cfg_file = get_config_path('counterparty_config.json')
    try:
        with open(cfg_file, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        safe_print(f'[Error] 保存往来单位配置失败: {e}')
        return False


def load_header_mapping():
    """从本地 JSON 文件读取表头映射配置"""
    cfg_file = get_config_path('header_mapping_config.json')
    if os.path.exists(cfg_file):
        try:
            with open(cfg_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            safe_print(f'[Warn] 读取表头映射配置失败: {e}')
    return {'summary': {}, 'detail': {}}


def save_header_mapping(mapping):
    """保存表头映射配置到本地 JSON"""
    cfg_file = get_config_path('header_mapping_config.json')
    try:
        with open(cfg_file, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        safe_print(f'[Error] 保存表头映射配置失败: {e}')
        return False


def auto_match_columns(headers, field_defs):
    '''
    根据表头文字与字段定义关键词，自动计算最佳匹配列索引 (0-indexed)
    返回: {field_key: col_idx}
    '''
    mapping = { }
    if not headers:
        return mapping
    for f_key, f_def in field_defs.items():
        for kw in f_def['keywords']:
            clean_kw = kw.strip().replace(' ', '').lower()
            found = False
            for idx, h in enumerate(headers):
                if h is None:
                    continue
                clean_h = str(h).strip().replace(' ', '').lower()
                if clean_kw == clean_h:
                    mapping[f_key] = idx
                    found = True
                    break
            if found:
                break
    for f_key, f_def in field_defs.items():
        if f_key in mapping:
            continue
        for kw in f_def['keywords']:
            clean_kw = kw.strip().replace(' ', '').lower()
            found = False
            for idx, h in enumerate(headers):
                if h is None:
                    continue
                clean_h = str(h).strip().replace(' ', '').lower()
                if clean_kw in clean_h or clean_h in clean_kw:
                    mapping[f_key] = idx
                    found = True
                    break
            if found:
                break
    return mapping


def format_date_str(date_val):
    '''标准化日期格式为 YYYYMMDD'''
    if not date_val:
        return datetime.datetime.now().strftime('%Y%m%d')
    if hasattr(date_val, 'strftime'):
        return date_val.strftime('%Y%m%d')
    s = str(date_val).split(' ')[0].strip()
    s_clean = s.replace('.', '').replace('-', '').replace('/', '').strip()
    if len(s_clean) == 8 and s_clean.isdigit():
        return s_clean
    parts = re.split('[\\.\\-\\/\\s]+', s)
    if len(parts) == 3:
        y = parts[0]
        m = parts[1].zfill(2)
        d = parts[2].zfill(2)
        if len(y) == 2:
            y = '20' + y
        return f'''{y}{m}{d}'''
    return s_clean


def safe_float(val, default = 0.0):
    if val is None or val == '':
        return default
    
    try:
        if isinstance(val, str):
            val = val.replace(',', '').replace('¥', '').replace('$', '').replace('RMB', '').strip()
        return float(val)
    except (ValueError, TypeError):
        return default



def safe_int(val, default = 0):
    if val is None or val == '':
        return default
    
    try:
        if isinstance(val, str):
            val = val.replace(',', '').strip()
        return int(float(val))
    except (ValueError, TypeError):
        return default



def get_row_val(row, col_idx, default = None):
    if col_idx is not None and 0 <= col_idx < len(row):
        val = row[col_idx]
        if val is not None:
            return val
        return default
    return default


def parse_source_shipment_excel(file_path, custom_mapping = None, log_func = None):
    '''
    智能解析发货记录汇总表（如：2026年发货记录汇总-坦桑尼亚.xlsx）
    支持用户自定义表头映射与动态列位置适配！
    返回: (containers_list)
    '''
    
    def log(msg):
        if log_func:
            log_func(msg)
            return None
        safe_print(msg)

    if not os.path.exists(file_path):
        raise FileNotFoundError(f'''未找到源文件: {file_path}''')
    log(f'''[*] 正在载入发货记录汇总工作簿: {os.path.basename(file_path)}...''')
    wb = openpyxl.load_workbook(file_path, read_only = True, data_only = True)
    summary_sheet_name = None
    detail_sheet_name = None
    for name in wb.sheetnames:
        clean_name = name.strip()
        if '汇总' in clean_name:
            summary_sheet_name = name
            continue
        if '明细' in clean_name or '发货' in clean_name:
            detail_sheet_name = name
    if not summary_sheet_name:
        summary_sheet_name = wb.sheetnames[0]
    if not detail_sheet_name:
        detail_sheet_name = wb.sheetnames[1] if len(wb.sheetnames) > 1 else wb.sheetnames[0]
    log(f'''[*] 锁定货柜汇总表: [{summary_sheet_name}], 明细数据表: [{detail_sheet_name}]''')
    ws_sum = wb[summary_sheet_name]
    ws_det = wb[detail_sheet_name]
    counterparty_map = load_counterparty_map()
    known_parties = list(counterparty_map.keys())
    sum_rows = []
    empty_cnt = 0
    for r in ws_sum.iter_rows(values_only = True):
        if any(r):
            sum_rows.append(r)
            empty_cnt = 0
            continue
        empty_cnt += 1
        if empty_cnt > 30 and len(sum_rows) > 5:
            break
    det_rows = []
    empty_cnt = 0
    for r in ws_det.iter_rows(values_only = True):
        if any(r):
            det_rows.append(r)
            empty_cnt = 0
            continue
        empty_cnt += 1
        if empty_cnt > 50 and len(det_rows) > 10:
            break
    sum_header_idx = -1
    for i, row in enumerate(sum_rows[:10]):
        row_str = ' '.join(( str(c or '') for c in row ))
        if '合同号' in row_str:
            if not (not ('集装箱' in row_str) and not ('海运费' in row_str) and not ('外贸公司' in row_str)) or '序号' in row_str:
                sum_header_idx = i
                break
    if sum_header_idx == -1:
        sum_header_idx = 2
    sum_headers = [str(c or '').strip() for c in sum_rows[sum_header_idx]]
    log(f'''[*] 汇总表头位置: 第 {sum_header_idx + 1} 行, 字段数: {len(sum_headers)}''')
    active_mapping = load_header_mapping()
    if custom_mapping and 'summary' in custom_mapping and custom_mapping['summary']:
        sum_map = custom_mapping['summary']
    elif active_mapping.get('summary'):
        sum_map = active_mapping['summary']
    else:
        sum_map = auto_match_columns(sum_headers, SUMMARY_FIELD_DEFS)
    mapped_sum_desc = ', '.join(( f'''{SUMMARY_FIELD_DEFS[k]['label']}:列{v + 1}''' for k, v in sum_map.items() if k in SUMMARY_FIELD_DEFS and v is not None ))
    log(f'''[*] 汇总表列映射: {mapped_sum_desc}''')
    containers = []
    for r_idx, r in enumerate(sum_rows[sum_header_idx + 1:], start = sum_header_idx + 2):
        if not r or not any(r):
            continue
        seq_no = len(containers) + 1
        col_seq = get_row_val(r, sum_map.get('seq_no'), seq_no)
        col_trade = str(get_row_val(r, sum_map.get('trade_co'), '') or '').strip()
        col_contract = str(get_row_val(r, sum_map.get('contract_no'), '') or '').strip()
        col_date = get_row_val(r, sum_map.get('date'), '')
        col_bl = str(get_row_val(r, sum_map.get('bl_no'), '') or '').strip()
        col_container = str(get_row_val(r, sum_map.get('container_no'), '') or '').strip()
        col_boxes = safe_int(get_row_val(r, sum_map.get('boxes'), 0))
        col_total_price = safe_float(get_row_val(r, sum_map.get('total_price'), 0.0))
        col_usd_total = safe_float(get_row_val(r, sum_map.get('usd_total'), 0.0))
        col_ocean = safe_float(get_row_val(r, sum_map.get('ocean_usd'), 0.0))
        col_truck = safe_float(get_row_val(r, sum_map.get('truck_rmb'), 0.0))
        col_insurance = safe_float(get_row_val(r, sum_map.get('insurance_usd'), 0.0))
        col_agency = safe_float(get_row_val(r, sum_map.get('agency_rmb'), 0.0))
        col_cbm = safe_float(get_row_val(r, sum_map.get('cbm'), 0.0))
        col_rate = safe_float(get_row_val(r, sum_map.get('rate'), 6.7322), 6.7322)
        if not col_trade:
            for cell_val in r:
                if cell_val:
                    s_cell = str(cell_val).strip()
                    for kp in known_parties:
                        if kp in s_cell:
                            col_trade = kp
                            break
                    if col_trade:
                        break
            if not col_trade:
                for cell_val in r:
                    if cell_val:
                        if '优势库' in str(cell_val):
                            col_trade = '优势库'
                            break
                        elif '颐顿' in str(cell_val):
                            col_trade = '颐顿'
                            break
        if not (col_contract and not (col_contract == 'None')) or not str(col_contract).strip():
            continue
        c_info = {
            'seq_no': safe_int(col_seq, seq_no),
            'trade_co': col_trade or '颐顿',
            'contract_no': col_contract,
            'container_no': col_container,
            'bl_no': col_bl,
            'date': format_date_str(col_date),
            'rate': col_rate if col_rate > 0 else 6.7322,
            'ocean_usd': col_ocean,
            'truck_rmb': col_truck,
            'insurance_usd': col_insurance,
            'agency_rmb': col_agency,
            'boxes': col_boxes,
            'cbm': col_cbm,
            'total_price': col_total_price,
            'usd_total': col_usd_total,
            'items': [] }
        containers.append(c_info)
    log(f'''[OK] 成功从汇总表解析出 {len(containers)} 个货柜基本信息''')
    det_header_idx = -1
    for i, row in enumerate(det_rows[:10]):
        row_str = ' '.join(( str(c or '') for c in row ))
        if '型号' in row_str:
            if not (not ('单价' in row_str) and not ('数量' in row_str)) or '品名' in row_str:
                det_header_idx = i
                break
    if det_header_idx == -1:
        det_header_idx = 2
    det_headers = [str(c or '').strip() for c in det_rows[det_header_idx]]
    log(f'''[*] 明细表头位置: 第 {det_header_idx + 1} 行, 字段数: {len(det_headers)}''')
    if custom_mapping and 'detail' in custom_mapping and custom_mapping['detail']:
        det_map = custom_mapping['detail']
    elif active_mapping.get('detail'):
        det_map = active_mapping['detail']
    else:
        det_map = auto_match_columns(det_headers, DETAIL_FIELD_DEFS)
    mapped_det_desc = ', '.join(( f'''{DETAIL_FIELD_DEFS[k]['label']}:列{v + 1}''' for k, v in det_map.items() if k in DETAIL_FIELD_DEFS and v is not None ))
    log(f'''[*] 明细表列映射: {mapped_det_desc}''')
    det_data_rows = det_rows[det_header_idx + 1:]
    log(f'''[*] 待处理有效明细行数: {len(det_data_rows)}''')
    for c in containers:
        c_contract = c['contract_no'].strip().upper()
        c_cont_no = c['container_no'].strip().upper()
        c_seq = c['seq_no']
        c_trade = str(c.get('trade_co', '') or '').strip()
        matched_items = []
        for r in det_data_rows:
            if not r or not any(r):
                continue
            r_seq = safe_int(get_row_val(r, det_map.get('seq_no')))
            r_contract = str(get_row_val(r, det_map.get('contract_no'), '') or '').strip().upper()
            r_cont_no = str(get_row_val(r, det_map.get('container_no'), '') or '').strip().upper()
            
            # 1. 基础合同号 / 柜号 / 序号校验
            shipment_match = False
            if r_contract and c_contract:
                if r_contract == c_contract or r_contract in c_contract or c_contract in r_contract:
                    shipment_match = True
            if not shipment_match and c_cont_no and r_cont_no:
                if r_cont_no == c_cont_no or r_cont_no in c_cont_no or c_cont_no in r_cont_no:
                    shipment_match = True
            if not shipment_match and r_seq is not None and c_seq is not None and r_seq == c_seq:
                shipment_match = True
            
            if not shipment_match:
                continue

            # 2. 往来单位 / 外贸公司归属区分（解决同一柜号/合同号下多家外贸公司拼柜问题）
            r_trade = str(get_row_val(r, det_map.get('trade_co'), '') or '').strip()
            if not r_trade:
                for cell_val in r:
                    if cell_val:
                        s_cell = str(cell_val).strip()
                        for kp in known_parties:
                            if kp in s_cell:
                                r_trade = kp
                                break
                        if r_trade:
                            break
                if not r_trade:
                    for cell_val in r:
                        if cell_val:
                            if '优势库' in str(cell_val):
                                r_trade = '优势库'
                                break
                            elif '颐顿' in str(cell_val):
                                r_trade = '颐顿'
                                break

            # 往来单位校验：若明细行与当前货柜均有外贸公司，必须相匹配
            if r_trade and c_trade:
                if not (c_trade in r_trade or r_trade in c_trade):
                    continue
            elif not r_trade and c_trade:
                # 明细无明确外贸公司时，若同一合同号存在多个不同往来单位（拼柜），优先依据序号校验
                same_contract_containers = [tc for tc in containers if tc['contract_no'].strip().upper() == c_contract]
                if len(same_contract_containers) > 1 and r_seq is not None and c_seq is not None and r_seq != c_seq:
                    continue

            is_match = True
            if is_match:
                cat = str(get_row_val(r, det_map.get('cat'), '') or '').strip()
                raw_model = str(get_row_val(r, det_map.get('model'), '') or '').strip()
                boxes = safe_int(get_row_val(r, det_map.get('boxes'), 0))
                pack_qty = safe_int(get_row_val(r, det_map.get('pack_qty'), 1), 1)
                qty = safe_int(get_row_val(r, det_map.get('qty')), boxes * pack_qty)
                unit_price = safe_float(get_row_val(r, det_map.get('unit_price'), 0.0))
                total_price = safe_float(get_row_val(r, det_map.get('total_price')), unit_price * qty)
                usd_total = safe_float(get_row_val(r, det_map.get('usd_total')), round(total_price / c['rate'], 2))
                untaxed_unit = safe_float(get_row_val(r, det_map.get('untaxed_unit')), round(unit_price / 1.13, 4))
                untaxed_total = safe_float(get_row_val(r, det_map.get('untaxed_total')), round(untaxed_unit * qty, 2))
                cbm = safe_float(get_row_val(r, det_map.get('cbm'), 0.0))
                if raw_model:
                    item_code = sanitize_item_code(raw_model)
                    matched_items.append({
                        'cat': cat or raw_model,
                        'raw_model': raw_model,
                        'model': item_code,
                        'boxes': boxes,
                        'pack_qty': pack_qty if pack_qty > 0 else 1,
                        'qty': qty,
                        'unit_price': unit_price,
                        'total_price': total_price,
                        'usd_total': usd_total,
                        'untaxed_unit': untaxed_unit,
                        'untaxed_total': untaxed_total,
                        'cbm': cbm })
        c['items'] = matched_items
        c_calc_cbm = sum(( x['cbm'] for x in matched_items ))
        if c['cbm'] == 0 or abs(c['cbm'] - c_calc_cbm) > 0.01:
            c['cbm'] = round(c_calc_cbm, 4)
        log(f'''  - 货柜 [{c['contract_no']} | {c['trade_co']}] 匹配到 {len(matched_items)} 条品目明细, 总体积: {c['cbm']:.2f} CBM''')
    wb.close()
    return containers


def load_master_from_excel(file_path):
    '''从已有的上传综合表或 ESA009M 导入品目编码主数据（自动清洗空格）'''
    if not os.path.exists(file_path):
        return list(DEFAULT_EXISTING_MASTER)
    wb = openpyxl.load_workbook(file_path, read_only = True, data_only = True)
    master_rows = []
    target_sheet = None
    for s in wb.sheetnames:
        if not (not ('品目' in s) and not ('主数据' in s) and not ('ESA009M' in s)) or 'Item' in s:
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]
    ws = wb[target_sheet]
    rows = list(ws.iter_rows(values_only = True))
    if len(rows) > 1:
        for r in rows[1:]:
            if r and any(r) and r[0]:
                raw_code = str(r[0]).strip()
                code = sanitize_item_code(raw_code)
                name = str(r[1] or code).strip()
                item_type = str(r[2] or '[商品]').strip()
                spec = str(r[3] or '').strip()
                comb1 = str(r[4] or '').strip()
                kw = str(r[5] or '').strip()
                use_yn = str(r[6] or 'YES').strip()
                conv_rate = r[7] if len(r) > 7 and r[7] is not None else 1
                pack_qty = r[8] if len(r) > 8 and r[8] is not None else ''
                cbm = r[9] if len(r) > 9 and r[9] is not None else ''
                weight = r[10] if len(r) > 10 and r[10] is not None else ''
                master_rows.append((code, name, item_type, spec, comb1, kw, use_yn, conv_rate, pack_qty, cbm, weight))
    wb.close()
    if master_rows:
        return master_rows
    return list(DEFAULT_EXISTING_MASTER)


def clone_template_header_and_styles(src_ws, dst_ws):
    '''
    克隆模板工作表的第 1 行表头、单元格样式、数字格式、列宽以及【Excel 批注 (Comments)】
    '''
    for col_idx in range(1, src_ws.max_column + 1):
        src_cell = src_ws.cell(row = 1, column = col_idx)
        dst_cell = dst_ws.cell(row = 1, column = col_idx, value = src_cell.value)
        if src_cell.has_style:
            dst_cell.font = copy.copy(src_cell.font)
            dst_cell.fill = copy.copy(src_cell.fill)
            dst_cell.border = copy.copy(src_cell.border)
            dst_cell.alignment = copy.copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format
        if src_cell.comment:
            dst_cell.comment = Comment(src_cell.comment.text, src_cell.comment.author)
    for col_letter, col_dim in src_ws.column_dimensions.items():
        if col_dim.width:
            dst_ws.column_dimensions[col_letter].width = col_dim.width
    if src_ws.row_dimensions[1].height:
        dst_ws.row_dimensions[1].height = src_ws.row_dimensions[1].height
    else:
        dst_ws.row_dimensions[1].height = 25
    headers = [ str(src_ws.cell(row = 1, column = c).value or '').strip() for c in range(1, src_ws.max_column + 1) ]
    return headers


def build_erp_workbook(container_info, items_data, existing_master = None, output_path = 'ECOUNT_ERP_全套上传表.xlsx', template_paths = None, export_individual_files = False, log_func = None):
    '''
    生成 ECOUNT ERP 全套标准上传工作簿（支持自动识别并应用上传格式模板中的列顺序、样式与批注！）
    '''
    
    def log(msg):
        if log_func:
            log_func(msg)
            return None
        safe_print(msg)

    if existing_master is None:
        existing_master = list(DEFAULT_EXISTING_MASTER)
    if template_paths is None:
        template_paths = load_template_config()
    counterparty_map = load_counterparty_map()
    trade_co = container_info.get('trade_co', '颐顿')
    co_map = counterparty_map.get(trade_co, {
        'rmb_code': container_info.get('rmb_code', '33108107'),
        'rmb_name': container_info.get('rmb_name', f'''浙江{trade_co}机电有限公司'''),
        'usd_code': container_info.get('usd_code', '3310810701'),
        'usd_name': container_info.get('usd_name', f'''浙江{trade_co}机电有限公司usd''') })
    rate = safe_float(container_info.get('rate', 6.7322), 6.7322)
    contract_no = str(container_info.get('contract_no', 'TS2026002')).strip()
    container_no = str(container_info.get('container_no', '')).strip()
    date_str = format_date_str(container_info.get('date', '20260815'))
    seq_no = safe_int(container_info.get('seq_no', 1), 1)
    warehouse = str(container_info.get('warehouse', '100')).strip() or '100'
    ocean_usd = safe_float(container_info.get('ocean_usd', 0.0))
    ocean_rmb = round(ocean_usd * rate, 2)
    truck_rmb = safe_float(container_info.get('truck_rmb', 0.0))
    insurance_usd = safe_float(container_info.get('insurance_usd', 0.0))
    insurance_rmb = round(insurance_usd * rate, 2)
    agency_rmb = safe_float(container_info.get('agency_rmb', 0.0))
    total_alloc_fees = round(ocean_rmb + truck_rmb + insurance_rmb + agency_rmb, 2)
    total_3fees = round(ocean_rmb + truck_rmb + insurance_rmb, 2)
    total_cbm = sum(( safe_float(x.get('cbm', 0)) for x in items_data ))
    fee_desc_label = '四费' if agency_rmb > 0 else '三费'
    log(f'''[*] 开始生成货柜 [{contract_no} / {trade_co}] 上传表...''')
    log(f'''    费用参数: 海运费 ${ocean_usd:.2f} (RMB {ocean_rmb:.2f}), 托箱费 RMB {truck_rmb:.2f}, 保费 ${insurance_usd:.2f} (RMB {insurance_rmb:.2f}), 代理费 RMB {agency_rmb:.2f}, 汇率 {rate}''')
    log(f'''    品目总数: {len(items_data)} 个型号, 总体积: {total_cbm:.4f} CBM, {fee_desc_label}附带费用总计: RMB {total_alloc_fees:.2f}''')
    processed_items = []
    total_untaxed = 0.0
    total_tax = 0.0
    total_qty = 0
    total_usd = 0.0
    for item in items_data:
        cbm = safe_float(item.get('cbm', 0))
        ratio = cbm / total_cbm if total_cbm > 0 else 0
        fee_alloc = round(total_alloc_fees * ratio, 2)
        qty = safe_int(item.get('qty', 1), 1)
        boxes = safe_int(item.get('boxes', 1), 1)
        pack_qty = safe_int(item.get('pack_qty', 1), 1)
        unit_price = safe_float(item.get('unit_price', 0.0))
        total_price = safe_float(item.get('total_price', unit_price * qty))
        untaxed_unit = safe_float(item.get('untaxed_unit', unit_price / 1.13))
        untaxed_total = safe_float(item.get('untaxed_total', round(untaxed_unit * qty, 2)))
        tax = round(untaxed_total * 0.13, 2)
        usd_total = safe_float(item.get('usd_total', round(total_price / rate, 2)))
        total_untaxed += untaxed_total
        total_tax += tax
        total_qty += qty
        total_usd += usd_total
        raw_m = str(item.get('raw_model', item.get('model', ''))).strip()
        clean_code = sanitize_item_code(item.get('model', raw_m))
        processed_items.append({
            'code': clean_code,
            'model': clean_code,
            'raw_model': raw_m,
            'cat': str(item.get('cat', raw_m)).strip(),
            'boxes': boxes,
            'pack_qty': pack_qty,
            'qty': qty,
            'cbm': cbm,
            'ratio': ratio,
            'unit_price': unit_price,
            'total_price': total_price,
            'untaxed_unit': untaxed_unit,
            'untaxed_total': untaxed_total,
            'tax': tax,
            'usd_total': usd_total,
            'fee_alloc': fee_alloc })
    alloc_sum = sum(( x['fee_alloc'] for x in processed_items ))
    diff = round(total_alloc_fees - alloc_sum, 2)
    if abs(diff) > 0.0001 and processed_items:
        best_idx = 0
        max_c = -1
        for idx_it, it in enumerate(processed_items):
            if it['cbm'] > max_c:
                max_c = it['cbm']
                best_idx = idx_it
        processed_items[best_idx]['fee_alloc'] = round(processed_items[best_idx]['fee_alloc'] + diff, 2)
    container_cost = round(total_untaxed + total_alloc_fees, 2)
    header_font = Font(name = '微软雅黑', size = 10, bold = True, color = 'FFFFFF')
    header_fill = PatternFill(start_color = '1F4E78', end_color = '1F4E78', fill_type = 'solid')
    total_fill = PatternFill(start_color = 'D9E1F2', end_color = 'D9E1F2', fill_type = 'solid')
    thin_border = Border(left = Side(style = 'thin', color = 'D3D3D3'), right = Side(style = 'thin', color = 'D3D3D3'), top = Side(style = 'thin', color = 'D3D3D3'), bottom = Side(style = 'thin', color = 'D3D3D3'))
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = '一般凭证上传格式'
    prefix = f'''合同号:{contract_no} 柜号:{container_no}''' if container_no else f'''合同号:{contract_no}'''
    date_info = f'''发出日期:{date_str}'''
    rate_info = f'''汇率:{rate}'''
    summary_truck = f'''{prefix} 内陆托箱费 人民币:¥{truck_rmb:.2f} {rate_info} {date_info}'''
    summary_agency = f'''{prefix} 代理费 人民币:¥{agency_rmb:.2f} {rate_info} {date_info}'''
    summary_tax = f'''{prefix} 货款增值税款 人民币:¥{total_tax:.2f} {rate_info} {date_info}'''
    summary_goods = f'''{prefix} 采购货款(未税) 美金:${total_usd:.2f} 人民币:¥{total_untaxed:.2f} {rate_info} {date_info}'''
    summary_ocean = f'''{prefix} 海运费 美金:${ocean_usd:.2f} 人民币:¥{ocean_rmb:.2f} {rate_info} {date_info}'''
    summary_insurance = f'''{prefix} 保费 美金:${insurance_usd:.2f} 人民币:¥{insurance_rmb:.2f} {rate_info} {date_info}'''
    summary_vat_in = f'''{prefix} 货柜进项税 人民币:¥{total_tax:.2f} {rate_info} {date_info}'''
    summary_cost = f'''{prefix} 货柜综合到岸总成本 人民币:¥{container_cost:.2f} {rate_info} {date_info}'''
    voucher_data_entries = [
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'amount': truck_rmb,
            'usd_amount': '',
            'rate': '',
            'summary': summary_truck,
            'dept': '00002' },
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'amount': agency_rmb,
            'usd_amount': '',
            'rate': '',
            'summary': summary_agency,
            'dept': '00002' },
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'amount': total_tax,
            'usd_amount': '',
            'rate': '',
            'summary': summary_tax,
            'dept': '00002' },
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'amount': total_untaxed,
            'usd_amount': '',
            'rate': '',
            'summary': summary_goods,
            'dept': '00002' },
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['usd_code'],
            'co_name': co_map['usd_name'],
            'amount': ocean_rmb,
            'usd_amount': ocean_usd,
            'rate': rate,
            'summary': summary_ocean,
            'dept': '00002' },
        {
            'dr_cr': 4,
            'account_code': '2202',
            'co_code': co_map['usd_code'],
            'co_name': co_map['usd_name'],
            'amount': insurance_rmb,
            'usd_amount': round(insurance_usd, 2),
            'rate': rate,
            'summary': summary_insurance,
            'dept': '00002' },
        {
            'dr_cr': 3,
            'account_code': '22210101',
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'amount': total_tax,
            'usd_amount': '',
            'rate': '',
            'summary': summary_vat_in,
            'dept': '00002' },
        {
            'dr_cr': 3,
            'account_code': '1402',
            'co_code': '',
            'co_name': '',
            'amount': container_cost,
            'usd_amount': '',
            'rate': '',
            'summary': summary_cost,
            'dept': '00002' }]
    v_tpl_path = template_paths.get('voucher')
    use_v_tpl = v_tpl_path and os.path.exists(v_tpl_path)
    if use_v_tpl:
        
        try:
            v_wb = openpyxl.load_workbook(v_tpl_path, data_only = True)
            v_src_ws = v_wb.active
            v_headers = clone_template_header_and_styles(v_src_ws, ws1)
            v_col_map = auto_match_columns(v_headers, VOUCHER_TEMPLATE_FIELD_DEFS)
            log(f'''[*] 已应用 [一般凭证] 模板: {os.path.basename(v_tpl_path)} ({len(v_headers)} 列, 带批注)''')
            v_wb.close()
        except Exception as e:
            log(f'''[Warn] 读取凭证模板失败 ({e})，采用标准格式''')
            use_v_tpl = False

    if not use_v_tpl:
        v_headers = [
            '凭证日期',
            '序号',
            '会计凭证No.',
            '类型',
            '科目编码',
            '往来单位编码',
            '往来单位名',
            '金额',
            '外币金额',
            '汇率',
            '摘要编码',
            '摘要',
            '部门',
            '项目记事板']
        ws1.row_dimensions[1].height = 25
        for col_idx, h in enumerate(v_headers, 1):
            c = ws1.cell(row = 1, column = col_idx, value = h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal = 'center', vertical = 'center')
            c.border = thin_border
        v_col_map = auto_match_columns(v_headers, VOUCHER_TEMPLATE_FIELD_DEFS)
    for r_idx, v_entry in enumerate(voucher_data_entries, 2):
        ws1.row_dimensions[r_idx].height = 22
        entry_val_map = {
            'date': date_str,
            'seq_no': seq_no,
            'voucher_no': '',
            'dr_cr': v_entry['dr_cr'],
            'account_code': v_entry['account_code'],
            'co_code': v_entry['co_code'],
            'co_name': v_entry['co_name'],
            'amount': v_entry['amount'],
            'usd_amount': v_entry['usd_amount'],
            'rate': v_entry['rate'],
            'summary_code': '',
            'summary': v_entry['summary'],
            'dept': v_entry['dept'],
            'project': '' }
        for c_idx in range(1, len(v_headers) + 1):
            cell = ws1.cell(row = r_idx, column = c_idx)
            cell.font = Font(name = '微软雅黑', size = 10)
            cell.border = thin_border
            matched_key = None
            for f_key, mapped_idx in v_col_map.items():
                if mapped_idx == c_idx - 1:
                    matched_key = f_key
                    break
            val = entry_val_map.get(matched_key, '') if matched_key else ''
            cell.value = val
            if matched_key in ('date', 'seq_no', 'voucher_no', 'dr_cr', 'account_code', 'co_code', 'dept', 'project'):
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                if matched_key in ('date', 'account_code', 'co_code', 'dept'):
                    cell.number_format = '@'
                continue
            if matched_key in ('co_name', 'summary'):
                cell.alignment = Alignment(horizontal = 'left', vertical = 'center')
                continue
            if matched_key in ('amount', 'usd_amount'):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                if val != '':
                    cell.number_format = '#,##0.00'
                continue
            if matched_key == 'rate':
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                if val != '':
                    cell.number_format = '0.0000'
    ws2 = wb.create_sheet(title = '采购入库格式')
    p_tpl_path = template_paths.get('purchase')
    use_p_tpl = p_tpl_path and os.path.exists(p_tpl_path)
    if use_p_tpl:
        
        try:
            p_wb = openpyxl.load_workbook(p_tpl_path, data_only = True)
            p_src_ws = p_wb.active
            p_headers = clone_template_header_and_styles(p_src_ws, ws2)
            p_col_map = auto_match_columns(p_headers, PURCHASE_TEMPLATE_FIELD_DEFS)
            log(f'''[*] 已应用 [采购入库] 模板: {os.path.basename(p_tpl_path)} ({len(p_headers)} 列, 带批注)''')
            p_wb.close()
        except Exception as e:
            log(f'''[Warn] 读取采购模板失败 ({e})，采用标准格式''')
            use_p_tpl = False

    if not use_p_tpl:
        p_headers = [
            '日期',
            '序号',
            '往来单位编码',
            '往来单位名',
            '职员',
            '接收仓库',
            '交易类型',
            '货币',
            '汇率',
            '合同号',
            '集装箱号码',
            '汇率',
            '发货日期',
            '品目编码',
            '品目名',
            '规格名',
            '数量',
            '单价(包括税)',
            '单价',
            '外币金额',
            '供应价',
            '附带费用',
            '增值税',
            '摘要',
            '序列号/批号']
        ws2.row_dimensions[1].height = 25
        for col_idx, h in enumerate(p_headers, 1):
            cell = ws2.cell(row = 1, column = col_idx, value = h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
            cell.border = thin_border
        p_col_map = auto_match_columns(p_headers, PURCHASE_TEMPLATE_FIELD_DEFS)
    for r_idx, item in enumerate(processed_items, 2):
        ws2.row_dimensions[r_idx].height = 20
        code = item['code']
        item_val_map = {
            'date': date_str,
            'seq_no': seq_no,
            'co_code': co_map['rmb_code'],
            'co_name': co_map['rmb_name'],
            'staff': '',
            'warehouse': warehouse,
            'trade_type': '20',
            'currency': '',
            'rate': rate,
            'contract_no': contract_no,
            'container_no': container_no,
            'rate_text': rate,
            'ship_date': date_str,
            'item_code': code,
            'item_name': code,
            'spec': item['raw_model'] or code,
            'qty': item['qty'],
            'tax_price': '',
            'untaxed_price': f'''货柜采购入库-{code}''',
            'usd_total': item['tax'],
            'untaxed_total': item['fee_alloc'],
            'fee_alloc': item['untaxed_total'],
            'tax': item['usd_total'],
            'remark': round(item['untaxed_unit'], 2),
            'batch_no': item['unit_price'] }
        for c_idx in range(1, len(p_headers) + 1):
            cell = ws2.cell(row = r_idx, column = c_idx)
            cell.font = Font(name = '微软雅黑', size = 10)
            cell.border = thin_border
            h_name = str(p_headers[c_idx - 1] or '').strip()
            matched_key = None
            for f_key, mapped_idx in p_col_map.items():
                if mapped_idx == c_idx - 1:
                    matched_key = f_key
                    break
            if h_name == '汇率' or matched_key in ('rate', 'rate_text'):
                val = rate
                matched_key = 'rate'
            elif '仓库' in h_name or matched_key == 'warehouse':
                val = warehouse
                matched_key = 'warehouse'
            elif matched_key:
                val = item_val_map.get(matched_key, '')
            else:
                val = ''
            cell.value = val
            if matched_key in ('date', 'seq_no', 'co_code', 'staff', 'warehouse', 'trade_type', 'currency', 'contract_no', 'container_no', 'ship_date', 'item_code', 'batch_no'):
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                if matched_key in ('date', 'co_code', 'trade_type', 'warehouse', 'contract_no', 'container_no', 'ship_date', 'item_code'):
                    cell.number_format = '@'
                continue
            if matched_key in ('co_name', 'item_name', 'spec', 'remark'):
                cell.alignment = Alignment(horizontal = 'left', vertical = 'center')
                if matched_key in ('item_name', 'spec'):
                    cell.number_format = '@'
                continue
            if matched_key in ('rate', 'rate_text'):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                cell.number_format = '0.0000'
                continue
            if matched_key == 'qty':
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                cell.number_format = '#,##0'
                continue
            if matched_key in ('tax_price', 'untaxed_price', 'usd_total', 'untaxed_total', 'fee_alloc', 'tax'):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                cell.number_format = '#,##0.00'
    tot_p_row = len(processed_items) + 2
    for c_idx in range(1, len(p_headers) + 1):
        cell = ws2.cell(row = tot_p_row, column = c_idx)
        cell.font = Font(name = '微软雅黑', size = 10, bold = True)
        cell.fill = total_fill
        cell.border = thin_border
        matched_key = None
        for f_key, mapped_idx in p_col_map.items():
            if mapped_idx == c_idx - 1:
                matched_key = f_key
                item_val_map.get(matched_key, '')
                break
        col_letter = get_column_letter(c_idx)
        if matched_key == 'item_name':
            cell.value = '合计'
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
            continue
        if matched_key == 'qty':
            cell.value = f'''=SUM({col_letter}2:{col_letter}{tot_p_row - 1})'''
            cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
            cell.number_format = '#,##0'
            continue
        if matched_key in ('usd_total', 'untaxed_total', 'fee_alloc', 'tax'):
            cell.value = f'''=SUM({col_letter}2:{col_letter}{tot_p_row - 1})'''
            cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
            cell.number_format = '#,##0.00'
    ws3 = wb.create_sheet(title = '新品目格式')
    n_tpl_path = template_paths.get('new_item')
    use_n_tpl = n_tpl_path and os.path.exists(n_tpl_path)
    if use_n_tpl:
        
        try:
            n_wb = openpyxl.load_workbook(n_tpl_path, data_only = True)
            n_src_ws = n_wb.active
            n_headers = clone_template_header_and_styles(n_src_ws, ws3)
            n_col_map = auto_match_columns(n_headers, NEW_ITEM_TEMPLATE_FIELD_DEFS)
            log(f'''[*] 已应用 [新增品目] 模板: {os.path.basename(n_tpl_path)} ({len(n_headers)} 列, 带批注)''')
            n_wb.close()
        except Exception as e:
            log(f'''[Warn] 读取新品目模板失败 ({e})，采用标准格式''')
            use_n_tpl = False

    if not use_n_tpl:
        n_headers = [
            '品目编码',
            '品目名',
            '规格类型',
            '规格',
            '单位',
            '品目类型',
            '使用套件',
            '数量管理',
            '入库单价',
            '入库单价 是否包括增值税',
            '生产流程',
            '出库单价',
            '出库单价 是否包括增值税',
            '品目组合1',
            '装数',
            '体积',
            '重量']
        ws3.row_dimensions[1].height = 25
        for col_idx, h in enumerate(n_headers, 1):
            cell = ws3.cell(row = 1, column = col_idx, value = h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
            cell.border = thin_border
        n_col_map = auto_match_columns(n_headers, NEW_ITEM_TEMPLATE_FIELD_DEFS)
    existing_codes = set(( sanitize_item_code(x[0]) for x in existing_master ))
    missing_items = [ x for x in processed_items if x['code'] not in existing_codes ]
    for r_idx, item in enumerate(missing_items, 2):
        ws3.row_dimensions[r_idx].height = 20
        unit_cbm = round(item['cbm'] / item['qty'], 6) if item['qty'] > 0 else 0
        new_val_map = {
            'item_code': item['code'],
            'item_name': item['code'],
            'spec_type': '1',
            'spec': item['raw_model'] or item['code'],
            'unit': '台',
            'item_type': '3',
            'is_kit': '0',
            'qty_mgmt': '1',
            'in_price': item['unit_price'],
            'in_price_tax': '1',
            'process': '',
            'out_price': 0,
            'out_price_tax': '',
            'comb1': '',
            'pack_qty': item['pack_qty'],
            'unit_cbm': unit_cbm,
            'unit_weight': '' }
        for c_idx in range(1, len(n_headers) + 1):
            cell = ws3.cell(row = r_idx, column = c_idx)
            cell.font = Font(name = '微软雅黑', size = 10)
            cell.border = thin_border
            matched_key = None
            for f_key, mapped_idx in n_col_map.items():
                if mapped_idx == c_idx - 1:
                    matched_key = f_key
                    break
            val = new_val_map.get(matched_key, '') if matched_key else ''
            cell.value = val
            if matched_key in ('item_code', 'spec_type', 'unit', 'item_type', 'is_kit', 'qty_mgmt', 'in_price_tax', 'process', 'out_price_tax', 'comb1'):
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                if matched_key in ('item_code', 'spec_type', 'item_type', 'is_kit', 'qty_mgmt', 'in_price_tax'):
                    cell.number_format = '@'
                continue
            if matched_key in ('item_name', 'spec'):
                cell.alignment = Alignment(horizontal = 'left', vertical = 'center')
                cell.number_format = '@'
                continue
            if matched_key in ('in_price', 'out_price'):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                cell.number_format = '#,##0.00'
                continue
            if matched_key == 'pack_qty':
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                cell.number_format = '#,##0'
                continue
            if matched_key in ('unit_cbm', 'unit_weight'):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                if val != '':
                    cell.number_format = '0.000000'
    ws4 = wb.create_sheet(title = '品目编码主数据')
    master_headers = [
        '品目编码',
        '品目名',
        '品目类型',
        '规格信息',
        '品目组合1名',
        '查询关键词',
        '使用',
        '单位转换比率(分母)',
        '装数',
        '体积',
        '重量']
    ws4.row_dimensions[1].height = 25
    for col_idx, h in enumerate(master_headers, 1):
        cell = ws4.cell(row = 1, column = col_idx, value = h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
        cell.border = thin_border
    curr_r = 2
    for row in existing_master:
        ws4.row_dimensions[curr_r].height = 20
        clean_row = list(row)
        clean_row[0] = sanitize_item_code(clean_row[0])
        for col_idx, val in enumerate(clean_row, 1):
            cell = ws4.cell(row = curr_r, column = col_idx, value = val)
            cell.font = Font(name = '微软雅黑', size = 10)
            cell.border = thin_border
            if col_idx in (1, 3, 7):
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                continue
            if col_idx in (2, 4, 5, 6):
                cell.alignment = Alignment(horizontal = 'left', vertical = 'center')
                continue
            if col_idx in (8, 9, 10, 11):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
        curr_r += 1
    for item in missing_items:
        ws4.row_dimensions[curr_r].height = 20
        unit_cbm = round(item['cbm'] / item['qty'], 6) if item['qty'] > 0 else 0
        row = (item['code'], item['code'], '[商品]', item['raw_model'] or item['code'], '', '', 'YES', 1, item['pack_qty'], unit_cbm, '')
        for col_idx, val in enumerate(row, 1):
            cell = ws4.cell(row = curr_r, column = col_idx, value = val)
            cell.font = Font(name = '微软雅黑', size = 10)
            cell.border = thin_border
            if col_idx in (1, 3, 7):
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                continue
            if col_idx in (2, 4, 5, 6):
                cell.alignment = Alignment(horizontal = 'left', vertical = 'center')
                continue
            if col_idx in (8, 9, 10, 11):
                cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
        curr_r += 1
    has_agency = agency_rmb > 0
    fee_type_str = '四费' if has_agency else '三费'
    sheet5_title = f'''货柜{fee_type_str}分摊明细表'''
    ws5 = wb.create_sheet(title = sheet5_title)
    if has_agency:
        ws5.merge_cells('A1:O1')
        ws5['A1'] = f'''货柜（{contract_no} / {container_no} / {trade_co}）四费体积占比分摊核算明细表'''
        ws5['A1'].font = Font(name = '微软雅黑', size = 12, bold = True, color = '1F4E78')
        ws5['A1'].alignment = Alignment(horizontal = 'center', vertical = 'center')
        ws5.row_dimensions[1].height = 30
        f_headers = [
            '品名',
            '品目编码/型号',
            '箱数',
            '装数',
            '出货数量(台)',
            '体积(CBM)',
            '体积占比',
            '分摊海运费(RMB)',
            '分摊内托费(RMB)',
            '分摊保费(RMB)',
            '分摊代理费(RMB)',
            '分摊四费总计(RMB)',
            '单台分摊四费(元/台)',
            '采购未税总额(RMB)',
            '综合到岸成本(RMB)']
        ws5.row_dimensions[3].height = 25
        for col_idx, h in enumerate(f_headers, 1):
            cell = ws5.cell(row = 3, column = col_idx, value = h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
            cell.border = thin_border
        start_f_row = 4
        end_f_row = start_f_row + len(processed_items) - 1
        tot_f_row = end_f_row + 1
        for r_idx, item in enumerate(processed_items, start_f_row):
            ws5.row_dimensions[r_idx].height = 20
            row_data = [
                item['cat'],
                item['code'],
                item['boxes'],
                item['pack_qty'],
                item['qty'],
                item['cbm'],
                f'''=F{r_idx}/$F${tot_f_row}''' if total_cbm > 0 else 0,
                f'''=G{r_idx}*({ocean_usd}*{rate})''',
                f'''=G{r_idx}*{truck_rmb}''',
                f'''=G{r_idx}*({insurance_usd}*{rate})''',
                f'''=G{r_idx}*{agency_rmb}''',
                f'''=H{r_idx}+I{r_idx}+J{r_idx}+K{r_idx}''',
                f'''=L{r_idx}/E{r_idx}''' if item['qty'] > 0 else 0,
                item['untaxed_total'],
                f'''=N{r_idx}+L{r_idx}''']
            for col_idx, val in enumerate(row_data, 1):
                cell = ws5.cell(row = r_idx, column = col_idx, value = val)
                cell.font = Font(name = '微软雅黑', size = 10)
                cell.border = thin_border
                if col_idx in (1, 2):
                    cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                    continue
                if col_idx in (3, 4, 5):
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '#,##0'
                    continue
                if col_idx == 6:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '#,##0.00'
                    continue
                if col_idx == 7:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '0.00%'
                    continue
                if col_idx in (8, 9, 10, 11, 12, 14, 15):
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '¥#,##0.00'
                    continue
                if col_idx == 13:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '¥#,##0.0000'
        ws5.cell(row = tot_f_row, column = 2, value = '合计')
        ws5.cell(row = tot_f_row, column = 3, value = f'''=SUM(C{start_f_row}:C{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 5, value = f'''=SUM(E{start_f_row}:E{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 6, value = f'''=SUM(F{start_f_row}:F{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 7, value = f'''=SUM(G{start_f_row}:G{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 8, value = f'''=SUM(H{start_f_row}:H{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 9, value = f'''=SUM(I{start_f_row}:I{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 10, value = f'''=SUM(J{start_f_row}:J{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 11, value = f'''=SUM(K{start_f_row}:K{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 12, value = f'''=SUM(L{start_f_row}:L{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 13, value = f'''=L{tot_f_row}/E{tot_f_row}''')
        ws5.cell(row = tot_f_row, column = 14, value = f'''=SUM(N{start_f_row}:N{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 15, value = f'''=SUM(O{start_f_row}:O{end_f_row})''')
        for c in range(1, 16):
            cell = ws5.cell(row = tot_f_row, column = c)
            cell.font = Font(name = '微软雅黑', size = 10, bold = True)
            cell.fill = total_fill
            cell.border = thin_border
            if c in (3, 5):
                cell.number_format = '#,##0'
                continue
            if c == 6:
                cell.number_format = '#,##0.00'
                continue
            if c == 7:
                cell.number_format = '0.00%'
                continue
            if c in (8, 9, 10, 11, 12, 14, 15):
                cell.number_format = '¥#,##0.00'
                continue
            if c == 13:
                cell.number_format = '¥#,##0.0000'
                continue
            if c == 2:
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
    else:
        ws5.merge_cells('A1:N1')
        ws5['A1'] = f'''货柜（{contract_no} / {container_no} / {trade_co}）三费体积占比分摊核算明细表'''
        ws5['A1'].font = Font(name = '微软雅黑', size = 12, bold = True, color = '1F4E78')
        ws5['A1'].alignment = Alignment(horizontal = 'center', vertical = 'center')
        ws5.row_dimensions[1].height = 30
        f_headers = [
            '品名',
            '品目编码/型号',
            '箱数',
            '装数',
            '出货数量(台)',
            '体积(CBM)',
            '体积占比',
            '分摊海运费(RMB)',
            '分摊内托费(RMB)',
            '分摊保费(RMB)',
            '分摊三费总计(RMB)',
            '单台分摊三费(元/台)',
            '采购未税总额(RMB)',
            '综合到岸成本(RMB)']
        ws5.row_dimensions[3].height = 25
        for col_idx, h in enumerate(f_headers, 1):
            cell = ws5.cell(row = 3, column = col_idx, value = h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
            cell.border = thin_border
        start_f_row = 4
        end_f_row = start_f_row + len(processed_items) - 1
        tot_f_row = end_f_row + 1
        for r_idx, item in enumerate(processed_items, start_f_row):
            ws5.row_dimensions[r_idx].height = 20
            row_data = [
                item['cat'],
                item['code'],
                item['boxes'],
                item['pack_qty'],
                item['qty'],
                item['cbm'],
                f'''=F{r_idx}/$F${tot_f_row}''' if total_cbm > 0 else 0,
                f'''=G{r_idx}*({ocean_usd}*{rate})''',
                f'''=G{r_idx}*{truck_rmb}''',
                f'''=G{r_idx}*({insurance_usd}*{rate})''',
                f'''=H{r_idx}+I{r_idx}+J{r_idx}''',
                f'''=K{r_idx}/E{r_idx}''' if item['qty'] > 0 else 0,
                item['untaxed_total'],
                f'''=M{r_idx}+K{r_idx}''']
            for col_idx, val in enumerate(row_data, 1):
                cell = ws5.cell(row = r_idx, column = col_idx, value = val)
                cell.font = Font(name = '微软雅黑', size = 10)
                cell.border = thin_border
                if col_idx in (1, 2):
                    cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
                    continue
                if col_idx in (3, 4, 5):
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '#,##0'
                    continue
                if col_idx == 6:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '#,##0.00'
                    continue
                if col_idx == 7:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '0.00%'
                    continue
                if col_idx in (8, 9, 10, 11, 13, 14):
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '¥#,##0.00'
                    continue
                if col_idx == 12:
                    cell.alignment = Alignment(horizontal = 'right', vertical = 'center')
                    cell.number_format = '¥#,##0.0000'
        ws5.cell(row = tot_f_row, column = 2, value = '合计')
        ws5.cell(row = tot_f_row, column = 3, value = f'''=SUM(C{start_f_row}:C{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 5, value = f'''=SUM(E{start_f_row}:E{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 6, value = f'''=SUM(F{start_f_row}:F{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 7, value = f'''=SUM(G{start_f_row}:G{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 8, value = f'''=SUM(H{start_f_row}:H{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 9, value = f'''=SUM(I{start_f_row}:I{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 10, value = f'''=SUM(J{start_f_row}:J{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 11, value = f'''=SUM(K{start_f_row}:K{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 12, value = f'''=K{tot_f_row}/E{tot_f_row}''')
        ws5.cell(row = tot_f_row, column = 13, value = f'''=SUM(M{start_f_row}:M{end_f_row})''')
        ws5.cell(row = tot_f_row, column = 14, value = f'''=SUM(N{start_f_row}:N{end_f_row})''')
        for c in range(1, 15):
            cell = ws5.cell(row = tot_f_row, column = c)
            cell.font = Font(name = '微软雅黑', size = 10, bold = True)
            cell.fill = total_fill
            cell.border = thin_border
            if c in (3, 5):
                cell.number_format = '#,##0'
                continue
            if c == 6:
                cell.number_format = '#,##0.00'
                continue
            if c == 7:
                cell.number_format = '0.00%'
                continue
            if c in (8, 9, 10, 11, 13, 14):
                cell.number_format = '¥#,##0.00'
                continue
            if c == 12:
                cell.number_format = '¥#,##0.0000'
                continue
            if c == 2:
                cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
    for ws in (ws1, ws2, ws3, ws4, ws5):
        for col in ws.columns:
            max_len = max(( len(str(cell.value or '')) for cell in col ))
            col_letter = get_column_letter(col[0].column)
            if col_letter not in ws.column_dimensions or not (ws.column_dimensions[col_letter].width):
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok = True)
    wb.save(output_path)
    log(f'''[OK] 成功生成并保存 ECOUNT ERP 全套标准上传工作簿: {output_path}''')
    exported_individual_paths = []
    if export_individual_files:
        p_single_wb = openpyxl.Workbook()
        p_single_ws = p_single_wb.active
        p_single_ws.title = '输入采购入库'
        if use_p_tpl:
            p_wb = openpyxl.load_workbook(p_tpl_path, data_only = True)
            clone_template_header_and_styles(p_wb.active, p_single_ws)
            p_wb.close()
        else:
            for col_idx, h in enumerate(p_headers, 1):
                c = p_single_ws.cell(row = 1, column = col_idx, value = h)
                c.font = header_font
                c.fill = header_fill
                c.border = thin_border
        for r_idx in range(2, len(processed_items) + 2):
            for c_idx in range(1, len(p_headers) + 1):
                p_single_ws.cell(row = r_idx, column = c_idx, value = ws2.cell(row = r_idx, column = c_idx).value)
        p_single_path = os.path.join(out_dir, f'''{contract_no}_{trade_co}_采购入库上传.xlsx''')
        p_single_wb.save(p_single_path)
        exported_individual_paths.append(p_single_path)
        log(f'''  -> 已单独生成采购入库上传单表: {os.path.basename(p_single_path)}''')
        if missing_items:
            n_single_wb = openpyxl.Workbook()
            n_single_ws = n_single_wb.active
            n_single_ws.title = '新增品目'
            if use_n_tpl:
                n_wb = openpyxl.load_workbook(n_tpl_path, data_only = True)
                clone_template_header_and_styles(n_wb.active, n_single_ws)
                n_wb.close()
            else:
                for col_idx, h in enumerate(n_headers, 1):
                    c = n_single_ws.cell(row = 1, column = col_idx, value = h)
                    c.font = header_font
                    c.fill = header_fill
                    c.border = thin_border
            for r_idx in range(2, len(missing_items) + 2):
                for c_idx in range(1, len(n_headers) + 1):
                    n_single_ws.cell(row = r_idx, column = c_idx, value = ws3.cell(row = r_idx, column = c_idx).value)
            n_single_path = os.path.join(out_dir, f'''{contract_no}_{trade_co}_新增品目上传.xlsx''')
            n_single_wb.save(n_single_path)
            exported_individual_paths.append(n_single_path)
            log(f'''  -> 已单独生成新增品目上传单表: {os.path.basename(n_single_path)}''')
        v_single_wb = openpyxl.Workbook()
        v_single_ws = v_single_wb.active
        v_single_ws.title = '一般凭证'
        if use_v_tpl:
            v_wb = openpyxl.load_workbook(v_tpl_path, data_only = True)
            clone_template_header_and_styles(v_wb.active, v_single_ws)
            v_wb.close()
        else:
            for col_idx, h in enumerate(v_headers, 1):
                c = v_single_ws.cell(row = 1, column = col_idx, value = h)
                c.font = header_font
                c.fill = header_fill
                c.border = thin_border
        for r_idx in range(2, len(voucher_data_entries) + 2):
            for c_idx in range(1, len(v_headers) + 1):
                v_single_ws.cell(row = r_idx, column = c_idx, value = ws1.cell(row = r_idx, column = c_idx).value)
        v_single_path = os.path.join(out_dir, f'''{contract_no}_{trade_co}_一般凭证上传.xlsx''')
        v_single_wb.save(v_single_path)
        exported_individual_paths.append(v_single_path)
        log(f'''  -> 已单独生成一般凭证上传单表: {os.path.basename(v_single_path)}''')
    return output_path


def batch_process_containers(source_file, output_dir = None, selected_indices = None, master_data = None, custom_mapping = None, template_paths = None, export_individual_files = False, progress_callback = None, log_callback = None):
    '''
    批量处理发货汇总表中的货柜，并生成各柜对应的上传表（支持模板自适应与批注注入）
    '''
    
    def log(msg):
        if log_callback:
            log_callback(msg)
            return None
        safe_print(msg)

    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(source_file)), '报表输出')
    os.makedirs(output_dir, exist_ok = True)
    containers = parse_source_shipment_excel(source_file, custom_mapping = custom_mapping, log_func = log)
    if not containers:
        log('[Warn] 未在源文件中找到任何有效货柜数据！')
        return []
    if master_data is None:
        master_data = list(DEFAULT_EXISTING_MASTER)
    if template_paths is None:
        template_paths = load_template_config()
    if selected_indices is not None:
        target_containers = [ containers[i] for i in selected_indices if 0 <= i < len(containers) ]
    else:
        target_containers = containers
    log(f'''[*] 开始批量生成任务: 共选中 {len(target_containers)} 个货柜''')
    generated_files = []
    for idx, c in enumerate(target_containers, 1):
        contract_no = c.get('contract_no', f'''C{idx}''')
        trade_co = c.get('trade_co', '公司')
        filename = f'''{contract_no}_{trade_co}_ECOUNT全套上传表.xlsx'''
        filename = re.sub('[\\\\/*?:"<>|]', '_', filename)
        out_file = os.path.join(output_dir, filename)
        build_erp_workbook(c, c.get('items', []), existing_master = master_data, output_path = out_file, template_paths = template_paths, export_individual_files = export_individual_files, log_func = log)
        generated_files.append(out_file)
        if progress_callback:
            progress_callback(idx, len(target_containers))
    log(f'''[完成] 批量处理完毕！共生成 {len(generated_files)} 份工作簿，存放在: {output_dir}''')
    return generated_files

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description = 'ECOUNT ERP 货柜三费分摊与导入表生成工具')
    parser.add_argument('--file', '-f', help = '发货汇总表 Excel 路径')
    parser.add_argument('--output', '-o', help = '输出目录或文件路径')
    parser.add_argument('--gui', '-g', action = 'store_true', help = '启动图形交互界面 (GUI)')
    args = parser.parse_args()
    if args.gui or len(sys.argv) == 1:
        
        try:
            import container_gui
            container_gui.main()
        except ImportError:
            from container_gui import main as gui_main
            gui_main()

    if args.file:
        batch_process_containers(args.file, output_dir = args.output)
    safe_print('请提供源文件路径 (--file) 或启动图形界面 (--gui)')
