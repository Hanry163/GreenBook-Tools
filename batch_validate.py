#!/usr/bin/env python3
"""批量验证VIN码有效性

用法:
  python batch_validate.py                         # 扫描 baidu_ocr_output 的VIN
  python batch_validate.py <file.json>             # 验证指定JSON文件
  python batch_validate.py --check <vin1> <vin2>   # 直接输入VIN
"""
import sys, os, json, glob

sys.path.insert(0, os.path.dirname(__file__))
from vin_validator import validate_vin

def load_vins_from_baidu_ocr():
    """从百度OCR输出中提取所有VIN"""
    base = r'C:\Users\hanry\Desktop\扫描测试\baidu_ocr_output\details'
    if not os.path.isdir(base):
        print(f'目录不存在: {base}')
        return []
    
    vins = []
    for fpath in sorted(glob.glob(os.path.join(base, '*.json'))):
        with open(fpath, encoding='utf-8') as f:
            data = json.load(f)
        fname = data.get('file', os.path.basename(fpath))
        fields = data.get('fields', {})
        if not fields:
            results = data.get('results', [])
            if results:
                fields = results[0] if isinstance(results[0], dict) else {}
        vin = fields.get('vin', '')
        if vin:
            vins.append((fname, vin))
    return vins


def main():
    vins = []
    
    if len(sys.argv) >= 3 and sys.argv[1] == '--check':
        # 命令行直接输入VIN
        vins = [(f'VIN-{i+1}', v) for i, v in enumerate(sys.argv[2:])]
    elif len(sys.argv) > 1:
        # 从JSON文件读取
        fpath = sys.argv[1]
        with open(fpath, encoding='utf-8') as f:
            data = json.load(f)
        fields = data.get('fields', {})
        vin = fields.get('vin', '')
        if vin:
            vins = [(os.path.basename(fpath), vin)]
    else:
        # 自动从百度OCR输出读取
        vins = load_vins_from_baidu_ocr()
    
    if not vins:
        print('未找到VIN数据')
        return
    
    ok = 0
    fail = 0
    
    print(f'{"文件名":<35} {"VIN码":<23} 结果')
    print('-' * 75)
    for name, vin in vins:
        r = validate_vin(vin)
        status = '通过' if r['valid'] else '失败'
        if r['valid']:
            ok += 1
        else:
            fail += 1
        reason = ''
        if not r['valid'] and r.get('details'):
            reason = f' (期望{r["calculated"]}≠实际{r["check_digit"]})'
        elif r['errors']:
            reason = f' ({r["errors"][0]})'
        print(f'{name:<35} {vin:<23} {status}{reason}')
    
    print('-' * 75)
    print(f'总计: {len(vins)}, 通过: {ok}, 失败: {fail}, 通过率: {ok/len(vins)*100:.1f}%')


if __name__ == '__main__':
    main()
