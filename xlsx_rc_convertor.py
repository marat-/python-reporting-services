# -*- coding: utf-8 -*-
"""
    R1C1-format formulae to A1-type converter.
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Notice: Only Reporting Services 2012 (or higher) is supporting export reports to
            xlsx-format.
"""

import re


def col2str(num):
    """ Converts column number to literal format (eg. 27 = 'AA') """
    inum = num
    res = ''

    while inum > 0:
        buf = (inum - 1) % 26
        res += chr(buf + 65)
        inum = int((inum - buf) / 26)

    return res


def col2int(colstr):
    """ Converts column literal to number (eg. 'AA' = 27) """
    # TODO: works for range('A', 'ZZ')
    res = 0
    for i, s in enumerate(colstr):
        res += (ord(s) - 64) + res * 26
    res -= ord(colstr[0]) - 64

    return res


def convert_rc_formula(formula, address):
    """ Converts R1C1-typed formula to A1-type """
    assert isinstance(formula, (str, unicode))
    assert isinstance(address, (str, unicode))

    # Convert cell's string-address to tuple like as (row, col)
    addr_re = re.compile(r'(?P<row>[A-Z]+)(?P<col>[0-9]+)')
    addr = addr_re.search(address)
    address = (
        addr.group('row'),
        addr.group('col')
    )

    # Get cell offsets from formula
    form_re = re.compile(r'(?P<row_offset>R\[?\-?\d*\]?)(?P<col_offset>C\[?\-?\d*\]?)')
    form = form_re.findall(formula)
    # Parse rows and columns
    replace_list = []
    part_re = re.compile(r'(?P<offset>[RC]\[\-?\d+\]?)|(?P<abs>[RC]\d?)')
    for part in form:
        res_part = []
        for i, item in enumerate(part):
            srch = part_re.search(item).groupdict()
            if srch.get('abs'):
                res_part.append(
                    srch.get('abs')[1:] or address[i]
                )
            elif srch.get('offset') and i == 0:
                res_part.append(
                    col2str(
                        int(srch.get('offset')[2:-1]) + col2int(address[i])
                    )
                )
            elif srch.get('offset'):
                res_part.append(
                    int(srch.get('offset')[2:-1]) + int(address[i])
                )
        # Write parts to list
        replace_list.append(
            (
                '{0}{1}'.format(*part),
                '{0}{1}'.format(*res_part),
            )
        )

    # Replace formula and return
    for repl in replace_list:
        formula = formula.replace(repl[0], repl[1])

    return formula


if __name__ == '__main__':
    print('R1C1 to A1 convertation sample:')

    address = 'AA22'
    formula = 'SUM(R[-1]C[0]:R[-5]C[2])*RC*SUM(R[-1]C[0]:R[-5]C[-10])'

    print(convert_rc_formula(formula, address))
