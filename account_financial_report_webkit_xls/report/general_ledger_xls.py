# -*- coding: utf-8 -*-
# Copyright 2009-2016 Noviat
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).
import xlwt
from datetime import datetime
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx as report_xls
#from openerp.addons.report_xlsx.utils import rowcol_to_cell
from openerp.addons.account_financial_report_webkit.report.general_ledger \
    import GeneralLedgerWebkit
from openerp.tools.translate import _
# import logging
# _logger = logging.getLogger(__name__)

_column_sizes = [
    ('date', 12),
    ('period', 12),
    ('move', 20),
    ('journal', 12),
    ('account_code', 12),
    ('partner', 30),
    ('ref', 30),
    ('label', 45),
    ('counterpart', 30),
    ('debit', 15),
    ('credit', 15),
    ('bal', 15),
    ('cumul_bal', 15),
    ('curr_bal', 15),
    ('curr_code', 7),
]

def cell(row, col, row_abs=False, col_abs=False):
    """
    Convert numeric row/col notation to an Excel cell
    reference string in A1 notation.
    """
    d = col // 26
    m = col % 26
    chr1 = ""    # Most significant character in AA1
    if row_abs:
        row_abs = '$'
    else:
        row_abs = ''
    if col_abs:
        col_abs = '$'
    else:
        col_abs = ''
    if d > 0:
        chr1 = chr(ord('A') + d - 1)
    chr2 = chr(ord('A') + m)
    # Zero index to 1-index
    return col_abs + chr1 + chr2 + row_abs + str(row + 1)


class report_xlsx_format:
    styles = {
        'xls_title': [('bold', True), ('font_size', 24)],
        'bold': [('bold', True)],
        'underline': [('underline', True)],
        'italic': [('italic', True)],
        #'fill': 'pattern: pattern solid, fore_color %s;' % _pfc,
        'fill_green': [('bg_color', '#CCFFCC')],
        'fill_grey': [('bg_color', '#808080')],
        'fill_yellow': [('bg_color', '#FFFFCC')],
        'borders_all': [('border', 1)],
        'center': [('align', 'center')],
        'right': [('align', 'right')],
    }

    def _get_style(self, wb, keys):
        if not isinstance(keys, (list, tuple)):
            keys = (keys, )
        return wb.add_format(dict(reduce(
            lambda x,y: x+y, [self.styles[k] for k in keys])))

    def xls_row_template(self, specs, wanted_list):
        """
        Returns a row template.

        Input :
        - 'wanted_list': list of Columns that will be returned in the
                         row_template
        - 'specs': list with Column Characteristics
            0: Column Name (from wanted_list)
            1: Column Colspan
            2: Column Size (unit = the width of the character ’0′
                            as it appears in the sheet’s default font)
            3: Column Type
            4: Column Data
            5: Column Formula (or 'None' for Data)
            6: Column Style
        """
        xls_types = {
            'bool': False,
            'date': None,
            'text': '',
            'number': 0,
        }
        r = []
        col = 0
        for w in wanted_list:
            found = False
            for s in specs:
                if s[0] == w:
                    found = True
                    s_len = len(s)
                    c = list(s[:5])
                    # set write_cell_func or formula
                    if s_len > 5 and s[5] is not None:
                        c.append({'formula': s[5]})
                    else:
                        c.append({
                            'write_cell_func': xls_types[c[3]]})
                    # Set custom cell style
                    if s_len > 6 and s[6] is not None:
                        c.append(s[6])
                    else:
                        c.append(None)
                    # Set cell formula
                    if s_len > 7 and s[7] is not None:
                        c.append(s[7])
                    else:
                        c.append(None)
                    r.append((col, c[1], c))
                    col += c[1]
                    break
            if not found:
                _logger.warn("report_xls.xls_row_template, "
                             "column '%s' not found in specs", w)
        return r

    def xls_write_row(self, ws, row_pos, row_data,
                      row_style=None, set_column_size=False):
        r = ws.row(row_pos)
        for col, size, spec in row_data:
            data = spec[4]
            formula = spec[5].get('formula') and \
                xlwt.Formula(spec[5]['formula']) or None
            style = spec[6] and spec[6] or row_style
            if not data:
                # if no data, use default values
                data = xls_types_default[spec[3]]
            if size != 1:
                if formula:
                    ws.write_merge(
                        row_pos, row_pos, col, col + size - 1, data, style)
                else:
                    ws.write_merge(
                        row_pos, row_pos, col, col + size - 1, data, style)
            else:
                if formula:
                    ws.write(row_pos, col, formula, style)
                else:
                    spec[5]['write_cell_func'](r, col, data, style)
            if set_column_size:
                ws.col(col).width = spec[2] * 256
        return row_pos + 1
    def addrow(self, sheet, row, data, cols=0, fmt=None, set_size=False):
        fmt = fmt or {}
        for col, val in enumerate(data):
            sheet.write(row, col, val, fmt)
        if fmt:
            while col < cols-1:
                col += 1
                sheet.write(row, col, '', fmt)

class AttrDict(dict):
    def __init__(self, *args, **kwargs):
        super(AttrDict, self).__init__(*args, **kwargs)
        self.__dict__ = self


class GeneralLedgerXls(report_xls, report_xlsx_format):
    column_sizes = [x[1] for x in _column_sizes]

    def generate_xlsx_report(self, wb, data, objects):
        wb.formats[0].set_font_size(10)
        ws = wb.add_worksheet(self.title[:31])
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_width_to_pages = 1
        row_pos = 0

        # set print header/footer
        #ws.header_str = self.xls_headers['standard']
        #ws.footer_str = self.xls_footers['standard']

        # cf. account_report_general_ledger.mako
        initial_balance_text = {'initial_balance': _('Computed'),
                                'opening_balance': _('Opening Entries'),
                                False: _('No')}

        # Title
        row_pos = 0
        cell_style = self._get_style(wb, 'xls_title')
        _p = AttrDict(self.parser_instance.localcontext)
        report_name = ' - '.join([_p.report_name.upper(),
                                 _p.company.partner_id.name,
                                 _p.company.currency_id.name])
        ws.write(row_pos, 0, report_name)

        for i, size in enumerate(self.column_sizes):
            ws.set_column(i, i, size)
        row_pos += 2

        # Header Table
        style_header1 = self._get_style(wb, ('bold', 'fill_green', 'center'))
        style_header2 = self._get_style(wb, ('center'))

        ws.merge_range(row_pos, 0, row_pos, 1, _('Chart of Account'), style_header1)
        ws.merge_range(row_pos+1, 0, row_pos+1, 1, _p.chart_account.name, style_header2)

        ws.write(row_pos, 2, _('Fiscal Year'), style_header1)
        ws.write(row_pos+1, 2, _p.fiscalyear.name if _p.fiscalyear else '-', style_header2)

        df = _('From') + ': %s ' + _('To') + ': %s'
        if _p.filter_form(data) == 'filter_date':
            dfh = _('Dates Filter')
            df = df % (_p.start_date or '', _p.stop_date or '')
        else:
            dfh =  _('Periods Filter')
            df = df % (_p.start_period and _p.start_period.name or '',
                       _p.stop_period and _p.stop_period.name or '')
        ws.merge_range(row_pos, 3, row_pos, 5, dfh, style_header1)
        ws.merge_range(row_pos+1, 3, row_pos+1, 5, df, style_header2)

        ws.write(row_pos, 6, _('Accounts Filter'), style_header1)
        text = _p.accounts(data) and ', '.join([
            account.code for account in _p.accounts(data)]) or _('All')
        ws.write(row_pos+1, 6, text, style_header2)

        ws.write(row_pos, 7, _('Target Moves'), style_header1)
        ws.write(row_pos+1, 7, _p.display_target_move(data), style_header2)

        ws.merge_range(row_pos, 8, row_pos, 9, _('Initial Balance'), style_header1)
        text = initial_balance_text[_p.initial_balance_mode]
        ws.merge_range(row_pos+1, 8, row_pos+1, 9, text, style_header2)
        row_pos += 1

        ws.freeze_panes(row_pos, 0)
        row_pos += 1

        # Column Title Row
        row_pos += 1
        c_title_cell_style = self._get_style(wb, ('bold'))

        # Column Header Row
        # cell_format = _xs['bold'] + _xs['fill'] + _xs['borders_all']
        # c_hdr_cell_style = self._get_style(wb, ('bold', 'borders_all'))
        # c_hdr_cell_style_right = self._get_style(wb, ('bold', 'borders_all', 'right'))
        # c_hdr_cell_style_center = self._get_style(wb, ('bold', 'borders_all', 'center'))
        # c_hdr_cell_style_decimal = c_hdr_cell_style_right

        # Column Initial Balance Row
        # c_init_cell_style = self._get_style(wb, ('italic', 'borders_all'))
        # c_init_cell_style_decimal = c_init_cell_style

        style_account = self._get_style(wb, ('bold'))
        style_labels = self._get_style(wb, ('bold', 'fill_yellow'))
        style_labels_r = self._get_style(wb, ('bold', 'fill_yellow', 'right'))
        style_initial_balance = self._get_style(wb, ('italic'))


        import pdb;pdb.set_trace()
        # cell styles for ledger lines
        for account in _p.objects:
            # Write account
            name = ' - '.join([account.code, account.name])
            ws.write(row_pos, 0, name, style_account)
            row_pos += 1

            # Write labels
            ws.write_row(row_pos, 0, [
                _('Date'), _('Period'), _('Entry'), _('Journal'), _('Account'),
                _('Partner'), _('Reference'), _('Label'), _('Counterpart')],
                style_labels)
            ws.write_row(row_pos, 10, [
                _('Debit'), _('Credit'), _('Filtered Bal.'), _('Cumul. Bal.')
                ], style_labels_r)
            row_pos += 1

            row_start = row_pos
            cumul_balance = cumul_balance_curr = 0

            # Write initial balance
            display_initial_balance = _p['init_balance'][account.id] and \
                (_p['init_balance'][account.id].get(
                    'debit', 0.0) != 0.0 or
                    _p['init_balance'][account.id].get('credit', 0.0) != 0.0)
            if not display_initial_balance:
                ws.write(row_pos, 8, _('Initial Balance'), style_initial_balance)
                init_balance = _p['init_balance'][account.id]
                cumul_balance += init_balance.get('init_balance') or 0.0
                cumul_balance_curr += init_balance.get('init_balance_currency') or 0.0
                row = [
                    init_balance.get('debit') or 0.0,
                    init_balance.get('credit') or 0.0,
                    '',
                    cumul_balance,
                    ]
                if _p.amount_currency(data):
                    row.append(cumul_balance_curr)
                ws.write_row(row_pos, 10, row, style_initial_balance)
                row_pos += 1

            # Write lines
            for line in _p['ledger_lines'][account.id]:
                label_elements = [line.get('lname') or '']
                if line.get('invoice_number'):
                    label_elements.append(
                        "(%s)" % (line['invoice_number'],))
                label = ' '.join(label_elements)
                cumul_balance += line.get('balance') or 0.0
                row = [
                    line['ldate'] and datetime.strptime(line['ldate'], '%Y-%m-%d') or '',
                    line.get('period_code') or '',
                    line.get('move_name') or '',
                    line.get('jcode') or '',
                    account.code,
                    line.get('partner_name') or '',
                    line.get('lref'),
                    label,
                    line.get('counterparts') or '',
                    line.get('debit', 0.0),
                    line.get('credit', 0.0),
                    line.get('balance', 0.0),
                    cumul_balance,
                    ]
                if _p.amount_currency(data):
                    cumul_balance_curr += line.get('amount_currency') or 0.0
                    row += [
                        line.get('amount_currency') or 0.0,
                        line.get('currency_code') or '',
                    ]
                ws.write_row(row_pos, 0, row)
                row_pos += 1

            # Write Sums
            row = [
                _('Cumulated Balance on Account'),
                '=SUM(%s:%s)' % (cell(row_start, 9), cell(row_pos-1, 9)),
                '=SUM(%s:%s)' % (cell(row_start, 10), cell(row_pos-1, 10)),
                '=SUM(%s:%s)' % (cell(row_start, 11), cell(row_pos-1, 11)),
                '=%s-%s' % (cell(row_pos, 9), cell(row_pos, 10)),
                ]
            if _p.amount_currency(data):
                row += [
                    cumul_balance_curr,
                    line.get('currency_code') or '',
                ]
            ws.write_row(row_pos, 9, row)
            row_pos += 2


GeneralLedgerXls('report.account.account_report_general_ledger_xls',
                 'account.account',
                 parser=GeneralLedgerWebkit)
