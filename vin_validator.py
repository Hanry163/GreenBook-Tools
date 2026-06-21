#!/usr/bin/env python3
"""
VIN 校验位验证工具
根据《车辆识别码VIN校验位计算方法》实现

用法:
  python vin_validator.py <VIN>         # 验证单个VIN
  python vin_validator.py               # 交互模式
"""

VIN_CHAR_MAP = {
    '0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8,
    'J': 1, 'K': 2, 'L': 3, 'M': 4, 'N': 5, 'P': 7, 'R': 9,
    'S': 2, 'T': 3, 'U': 4, 'V': 5, 'W': 6, 'X': 7, 'Y': 8, 'Z': 9
}

WEIGHTS = [8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2]
INVALID_CHARS = {'I', 'O', 'Q'}


def clean_vin(vin):
    return vin.strip().upper()


def validate_vin(vin):
    result = {'valid': False, 'check_digit': '', 'calculated': '', 'errors': []}
    v = clean_vin(vin)

    if len(v) != 17:
        result['errors'].append(f'VIN长度应为17位，当前为{len(v)}位')
        return result

    for c in v:
        if c in INVALID_CHARS:
            result['errors'].append(f'包含非法字符"{c}"（I/O/Q不允许出现在VIN中）')
            return result
        if c not in VIN_CHAR_MAP:
            result['errors'].append(f'包含无法映射的字符"{c}"')
            return result

    result['check_digit'] = v[8]
    total = 0
    details = []
    for i, ch in enumerate(v):
        val = VIN_CHAR_MAP[ch]
        w = WEIGHTS[i]
        prod = val * w
        total += prod
        details.append((i + 1, ch, val, w, prod))

    rem = total % 11
    calculated = str(rem) if rem < 10 else 'X'
    result['calculated'] = calculated
    result['valid'] = (calculated == result['check_digit'])
    result['total'] = total
    result['remainder'] = rem
    result['details'] = details

    if not result['valid']:
        result['errors'].append(f'校验位不匹配：期望"{calculated}"，实际"{result["check_digit"]}"')

    return result


def format_result(vin, r):
    v = clean_vin(vin)
    lines = [f'VIN码: {v}', f'长度: {len(v)} 位', '']

    if not r['valid'] and not r.get('details'):
        lines.append('验证失败:')
        for e in r['errors']:
            lines.append(f'  - {e}')
        return '\n'.join(lines)

    lines.append('位置 | 字符 | 映射值 | 权重 | 乘积')
    lines.append('-' * 35)
    for pos, ch, val, weight, prod in r['details']:
        marker = '  <- 校验位' if pos == 9 else ''
        lines.append(f'  {pos:2d}  |  {ch}   |   {val}   |  {weight:2d}  |  {prod:3d}{marker}')

    lines.append('')
    lines.append(f'总和: {r["total"]}')
    lines.append(f'总和 / 11 = {r["total"] // 11} 余 {r["remainder"]}')
    lines.append(f'计算校验值: {r["calculated"]}')
    lines.append(f'实际第9位:  {r["check_digit"]}')
    lines.append('')
    lines.append('通过!' if r['valid'] else '失败!')
    if r['errors']:
        for e in r['errors']:
            lines.append(f'  - {e}')
    return '\n'.join(lines)


def main():
    import sys
    if len(sys.argv) > 1:
        for vin in sys.argv[1:]:
            r = validate_vin(vin)
            print(format_result(vin, r))
            print()
    else:
        print('VIN校验位验证工具 (输入 q 退出)')
        while True:
            vin = input('VIN: ').strip()
            if vin.lower() in ('q', 'quit', 'exit'):
                break
            if not vin:
                continue
            r = validate_vin(vin)
            print()
            print(format_result(vin, r))
            print()


if __name__ == '__main__':
    main()
