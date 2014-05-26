# -*- coding: utf-8 -*-
"""
    R1C1-format formulae to A1-type converter.
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Notice: Only Reporting Services 2012 (or higher) is supporting export reports to
            xlsx-format.
"""

import re


def col2str(num, run=0):
    """ Converts column number to literal format (eg. 27 = 'AA') """
    if run:
        inum = num
        res = ''

        while inum > 0:
            buf = (inum - 1) % 26
            res += chr(buf + 65)
            inum = int((inum - buf) / 26)
        res = res[::-1]
    else:
        res = num

    return res


def col2int(colstr):
    """ Converts column literal to number (eg. 'AA' = 27, 'AAA' = 703 etc.) """
    res = 0
    for i, s in enumerate(colstr.upper()[::-1]):
        res += (ord(s)-64)*(26**i)

    return res


def check_range(value, mode=0):
    """ Validate ranges of column (mode=1) or row (mode=0) """
    if not isinstance(value, int):
        value = int(value)
    bound = 16384 if mode else 1048756
    value %= bound

    return value


def convert_rc_formula(formula, address):
    """ Converts R1C1-typed formula to A1-type """
    assert isinstance(formula, (str, unicode))
    assert isinstance(address, (str, unicode))

    formula = formula.replace(';', ',')
    # Excluding for formulae with sheet links
    if '!' not in formula:
        formula = formula.upper().replace(' ', '')

    # Delete format from formula
    format_re = re.compile(r'@.*@')
    formula = re.sub(format_re, '', formula)

    address = address.upper()

    # Convert cell's string-address to tuple like as (row, col)
    addr_re = re.compile(r'(?P<col>[A-Z]+)(?P<row>[0-9]+)')
    addr = addr_re.search(address)
    address = (
        int(
            addr.group('row')
        ),
        col2int(
            addr.group('col')
        ),
    )

    # Get cell offsets from formula
    # TODO: '=R[1]C-RC1'
    form_re = re.compile(r'(?P<row_offset>R\[?\-?\d*]?)(?P<col_offset>C\[?\-?\d*]?)')
    form = form_re.findall(formula)
    # Parse rows and columns
    replace_list = []
    part_re = re.compile(r'(?P<offset>[RC]\[\-?\d+]?)|(?P<abs>[RC]\d*)')
    for part in form:
        res_part = []
        for i, item in enumerate(part):
            srch = part_re.search(item).groupdict()
            if srch.get('abs'):
                res_part.append(
                    col2str(
                        check_range(
                            srch.get('abs')[1:] or address[i],
                            mode=i
                        ),
                        run=i
                    )
                )
            elif srch.get('offset'):
                try:
                    res_part.append(
                        col2str(
                            check_range(
                                int(srch.get('offset')[2:-1]) + address[i],
                                mode=i
                            ),
                            run=i
                        )
                    )
                except:
                    raise BaseException
                    exit(1)


        # Write parts to list
        replace_list.append(
            (
                '{0}{1}'.format(*part),
                '{1}{0}'.format(*res_part),
            )
        )

    # Replace formula and return
    for repl in replace_list:
        formula = formula.replace(repl[0], repl[1], 1)

    return formula


def get_cell_format(formula):
    """ Get format from formula to set to formula's cell """
    assert isinstance(formula, (str, unicode))
    format_reg = re.compile(r'(?P<format>@.*@)')
    fmt_search = format_reg.search(formula)
    ex_format = fmt_search.group('format') if fmt_search else ''

    # Return string without @ and '
    return ex_format[1:-1]


if __name__ == '__main__':
    print('R1C1 to A1 convert sample:')

    address = 'G20'
    formula = "=R[-1]C-R[-1]C[1]"

    print(get_cell_format(formula))

    print(convert_rc_formula(formula, address))