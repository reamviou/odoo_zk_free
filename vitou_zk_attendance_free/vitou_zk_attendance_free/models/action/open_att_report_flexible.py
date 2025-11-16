# -*- coding: utf-8 -*-
###############################################################################
#
#    Copyright (C) 2024-TODAY,
#    Author: REAM Vitou (reamvitou@yahoo.com)
#    Tel: +855 17 82 66 82
#
###############################################################################

from odoo import models, fields, api, _
from odoo.exceptions import ValidationError, UserError

import datetime
from ..lic.globals import func_mod
from datetime import datetime, timedelta, time
from collections import defaultdict
import pytz

import numpy as np

#add***
import io
import json
from odoo.tools import json_default
try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter
#end add


try:
    from zk import ZK, const
except ImportError as e:
    raise ValidationError(_(e.message))


class VitouzkDownloadAttFlexible(models.TransientModel):
    _name = 'vitouzkf.open.report.flexible'
    _description = 'Open Attendance Monthly flexible'
    # _inherit = ['mail.thread']
    # _rec_name = "provider"
    # _sql_constraints = [
    #      ('name_unique', 'unique(name)', "Currency Code is duplicated!"),
    # ]

    date_from = fields.Date(string="Date From", default=fields.Date.today())
    date_to = fields.Date(string="Date To", default=fields.Date.today())
    department_id = fields.Many2one(comodel_name='hr.department', string='Department')

    # date_to = fields.Date(string="Date To", default=fields.Date.today())

    def action_open_report(self):
        """Button action for creating Sale Order Pdf Report"""

        data_get = self.get_data()
        data = {
            'data': data_get
        }
        # data = {
        #     'data_att': data_gen,
        #     'month': month
        # }
        # return ''
        # print(',=', data_get)

        if len(data_get)>0:
            return self.env.ref(
                'vitou_zk_attendance_free.report_template_vitouzkf_att_flexible_report').report_action(self, data=data)
        else:
            return self.info('Nothing to generate report', 'warning')

    # add ***

    def action_open_report_excel(self):


        data_get = self.get_data()
        data = {
            'data': data_get
        }
        # print(self.model_name)
        r = {
            'type': 'ir.actions.report',
            # "url": f"/report/xlsx_reports_att",
            'data': {'model': self._name,
                     'options': json.dumps(data, default=json_default),
                     # json.dumps(data, default=date_utils.json_default),
                     'output_format': 'xlsx',
                     'report_name': 'Attendance Report',
                     },
            'report_type': 'att_xlsx',
        }
        # print(r)
        return r

    # model_name = 'vitouzkf.attendance.all'

    # add ***
    def action_open_report_excel(self):
        data_get = self.get_data()
        data = {
            'data': data_get
        }
        # print(self.model_name)
        r = {
            'type': 'ir.actions.report',
            # "url": f"/report/xlsx_reports_att",
            'data': {'model': self._name,
                     'options': json.dumps(data, default=json_default),
                     # json.dumps(data, default=date_utils.json_default),
                     'output_format': 'xlsx',
                     'report_name': 'Attendance Report',
                     },
            'report_type': 'att_xlsx',
        }
        # print(r)
        return r

    model_name = 'vitouzkf.attendance.all.flexible'

    @api.model
    def get_data(self):
        # print('donwload...')
        att_mod = self.env['vitouzkf.attendance.all.flexible']
        hr_mod = self.env['hr.attendance']

        for rec in self:
            data = []
            dep_name_array = []
            date_from = rec.date_from
            date_to = rec.date_to

            d_from = date_from.strftime('%Y-%m-%d')
            d_to = date_to.strftime('%Y-%m-%d') + ' 23:59:59'

            if rec.department_id:
                domain_dep = ('department_id', '=', rec.department_id.id)
                domain = [('state', 'in', ['posted', 'done']), ('is_att', '=', True),
                          ('date_inout', '>=', d_from), ('date_inout', '<=', d_to),
                          domain_dep]
            else:
                domain = [('state', 'in', ['posted', 'done']), ('is_att', '=', True),
                          ('date_inout', '>=', d_from), ('date_inout', '<=', d_to)]

            # print(domain)

            att_get = att_mod.sudo().search(domain, order='employee_id asc')

            if not att_get:
                raise ValidationError('No data for these selections!')
                # self.info('No data for this month', 'warning'))

            emp_name = att_mod.read_group(
                domain=domain,
                fields=['employee_name', 'total_work_hour:sum'],
                groupby=['employee_name'],
                # order='employee_name asc'
            )

            dep_group = att_mod.read_group(
                domain=domain,
                fields=['department_id'],
                groupby=['department_id'],
            )

            for d in dep_group:
                dep_id = d["department_id"]
                dep_name = d["department_id"][1]
                dep_name_array.append({'dep_name': dep_name})

            # print('dg=', dep_name_array)

            # print('en==', emp_name)
            # emp_name_new = [{'employee_name': x['employee_name'],'total_work_hour':x['total_work_hour']} for x in emp_name]

            data_new = []
            for emp_n in emp_name:
                emp_name_get = emp_n['employee_name']
                emp_dept = self.env[func_mod].get_departmentname(emp_name_get)

                if rec.department_id:
                    domain_dep = ('department_id', '=', rec.department_id.id)
                    # for emp in emp_name_new: 'ilike' ,'-%'
                    domain_g = [('state', 'in', ['posted', 'done']), ('is_att', '=', True),
                                ('date_inout', '>=', d_from), ('date_inout', '<=', d_to),
                                ('employee_name', '=', emp_name_get), domain_dep]
                else:
                    domain_g = [('state', 'in', ['posted', 'done']), ('is_att', '=', True),
                                ('date_inout', '>=', d_from), ('date_inout', '<=', d_to),
                                ('employee_name', '=', emp_name_get)]

                att_groupby = att_mod.read_group(
                    domain=domain_g,
                    fields=['day', 'total_work_hour:sum'],
                    groupby=['day'],
                    orderby='day asc'
                )
                att_day_hour = [
                    {'day': 1, 'hour': 0},
                    {'day': 2, 'hour': 0},
                    {'day': 3, 'hour': 0},
                    {'day': 4, 'hour': 0},
                    {'day': 5, 'hour': 0},
                    {'day': 6, 'hour': 0},
                    {'day': 7, 'hour': 0},
                    {'day': 8, 'hour': 0},
                    {'day': 9, 'hour': 0},
                    {'day': 10, 'hour': 0},

                    {'day': 11, 'hour': 0},
                    {'day': 12, 'hour': 0},
                    {'day': 13, 'hour': 0},
                    {'day': 14, 'hour': 0},
                    {'day': 15, 'hour': 0},
                    {'day': 16, 'hour': 0},
                    {'day': 17, 'hour': 0},
                    {'day': 18, 'hour': 0},
                    {'day': 19, 'hour': 0},
                    {'day': 20, 'hour': 0},

                    {'day': 21, 'hour': 0},
                    {'day': 22, 'hour': 0},
                    {'day': 23, 'hour': 0},
                    {'day': 24, 'hour': 0},
                    {'day': 25, 'hour': 0},
                    {'day': 26, 'hour': 0},
                    {'day': 27, 'hour': 0},
                    {'day': 28, 'hour': 0},
                    {'day': 29, 'hour': 0},
                    {'day': 30, 'hour': 0},

                    {'day': 31, 'hour': 0},
                ]
                # print('a==', att_day_hour)

                # count = 0
                day = 0
                for att in att_groupby:
                    # print('==', att['day'],',')
                    day = att['day']
                    hour = att['total_work_hour']
                    indx = int(day) - 1
                    if indx >= 0:
                        att_day_hour[indx] = {
                            'day': day,
                            'hour': round(hour, 1)
                        }
                    # count = count+1
                    # print('d=', day, ', i=', indx)
                # print(att_day_hour)

                data_new.append({
                    'employee_name': emp_name_get,
                    'dep_name': emp_dept,
                    'total_work_hour': round(emp_n['total_work_hour'], 1),
                    'day_hour': att_day_hour

                })
        month = d_from + ' and ' + rec.date_to.strftime('%Y-%m-%d')
        data.append({
            'month': month,
            'dep_name': dep_name_array,
            'data': data_new

        })

        # print('xx=', data)

        return data

        # add ***

        # @api.model
    def get_xlsx_report_att(self, data, response):
        """Organizing xlsx report"""
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet()
        cell_format = workbook.add_format(
            {'font_size': '14px', 'bold': True, 'align': 'center',
             'border': True})
        head = workbook.add_format(
            {'align': 'center', 'bold': True, 'font_size': '23px',
             'border': True})
        body = workbook.add_format(
            {'align': 'left', 'text_wrap': True, 'border': True})
        sheet.merge_range('A1:F1', 'Attendance Report', head)
        # sheet.set_column('A2:F2', 18)
        # sheet.set_row(0, 30)
        # sheet.set_row(1, 20)
        sheet.write('A3', 'Staff Name', cell_format)
        sheet.write('B3', 'Department', cell_format)
        sheet.write('C3', 'Total Hours', cell_format)
        sheet.write('D3', 'day1', cell_format)
        sheet.write('E3', 'day2', cell_format)
        sheet.write('F3', 'day3', cell_format)
        sheet.write('G3', 'day4', cell_format)
        sheet.write('H3', 'day5', cell_format)
        sheet.write('I3', 'day6', cell_format)
        sheet.write('J3', 'day7', cell_format)
        sheet.write('K3', 'day8', cell_format)
        sheet.write('L3', 'day9', cell_format)
        sheet.write('M3', 'day10', cell_format)
        sheet.write('N3', 'day11', cell_format)
        sheet.write('O3', 'day12', cell_format)
        sheet.write('P3', 'day13', cell_format)
        sheet.write('Q3', 'day14', cell_format)
        sheet.write('R3', 'day15', cell_format)
        sheet.write('S3', 'day16', cell_format)
        sheet.write('T3', 'day17', cell_format)
        sheet.write('U3', 'day18', cell_format)
        sheet.write('V3', 'day19', cell_format)
        sheet.write('W3', 'day20', cell_format)
        sheet.write('X3', 'day21', cell_format)
        sheet.write('Y3', 'day22', cell_format)
        sheet.write('Z3', 'day23', cell_format)
        sheet.write('AA3', 'day24', cell_format)
        sheet.write('AB3', 'day25', cell_format)
        sheet.write('AC3', 'day26', cell_format)
        sheet.write('AD3', 'day27', cell_format)
        sheet.write('AE3', 'day28', cell_format)
        sheet.write('AF3', 'day29', cell_format)
        sheet.write('AG3', 'day30', cell_format)
        sheet.write('AH3', 'day31', cell_format)

        # 'name', 'status','slip_paid' 'slip_closed_shift''slip_closed_day''price_subtotal''price_tax''price_total'
        # 'opened_uid''opened_uname''opened_staff''closed_uid''closed_uname''closed_staff' 'date_from''date_to''report_type'

        row = 3
        column = 0
        # print(data)

        for i in data['data']:
            sheet.write('A2', 'Date Between ' + i['month'], body)
            for j in i['data']:
                sheet.write(row, 0, j['employee_name'], body)
                sheet.write(row, 1, j['dep_name'], body)
                sheet.write(row, 2, j['total_work_hour'], body)
                col_i = 3
                for x in j['day_hour']:
                    # print('d=', x['day'],',')
                    hour = x['hour'] or 0.0
                    sheet.write(row, col_i, hour, body)
                    # sheet.write(row, str(col_i), "{:.2f}".format(hour),
                    #         body)

                    col_i = col_i + 1

                row = row + 1
            # value = value + 1
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()

    def info(self, info, type):
        return self.env[func_mod].myinfo(info, type)

    def action_reload(self):
        self.env[func_mod].reload()

    def action_close(self):
        self.env[func_mod].close_popup()





