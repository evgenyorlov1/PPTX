#!/usr/bin/env python

from __future__ import unicode_literals

import copy
import math
import sys
from os import path

from lxml import etree

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, CMDHandler


class Dashboard2(PPTXGenerator):

    sheet_name = 'Liquidity Dashboard'
    template_shape_names = (
        'firm-chart-label',
        'liquidity-metrics-label',
        'estimated-uses-chart-label',
    )
    separate_charts = {
        'separate-headline-metrics': '2-Liquidity-Headline-Metrics.pptx',
        'separate-receivables-firm': '2-RECEIVABLE-FIRM.pptx',
        'separate-receivables-businesses': '2-RECEIVABLES-FUNCTION.pptx',
        'separate-key-metrics': '2-KEY-METRICS.pptx',
        'separate-cash-next-period': '2-CASH-NEXT-PERIOD.pptx',
    }

    #Get values from excel sheet to fill template OK
    def get_simple_fillers(self):
        # Up to $999,000,000 display as Millions of Dollars, with M (no decimal)
        # Greater than $1,000,000,000 display as Billions of Dollars, with B (2 decimal places)  
        def big_money(val):
            val = round(val/10.**6, 0)

            if val <= 999:
                return '$'+self.format_float(val, prec=0)+'M'
            else:
                return '$'+self.format_float(val/1000, prec=2, strip=False)+'B'

        # Get values from excel sheet to fill template.
        return {
            lambda v: self.format_float(v, prec=1, strip=False): [
                C('C4'),
            ],
            lambda v: self.format_float(abs(v), prec=1, strip=False)+' Weeks': [
                C('C5'),
            ],
            lambda v: self.format_float(min(abs(v), 99999), prec=0)+'BPS': [
                C('F6'), C('I6'), C('L6'),
            ],
            lambda v: self.format_percent(min(v, .999), prec=1, strip=False)+'%': [
                C('F5'), C('I5'), C('L5'),
            ],
            big_money: [
                C('F4'), C('I4'), C('L4'),
            ],
            lambda v: big_money(min(v, 99.99*1000)*10**6): [
                C('B25', 'G26'),
            ],
            lambda v: v: [
                C('A10', 'A12'),
                C('A18', 'A20'),
            ],
            lambda v: self.format_float(v, prec=1, strip=False): [
                C('B18', 'E20'),
            ],
            lambda v: '999' if v >= 999 else self.format_float(v, prec=1, strip=False): [
                C('H25', 'H26'),
            ],
            lambda v: '999' if v >= 999 else self.format_float(v, prec=0): [
                C('I25', 'I26'),
            ],
            lambda v: '$'+self.with_comma(v/10.**6): [
                C('C39', 'C41'),
            ],
        }

    #Fill HEADLINE METRICS OK
    def fill_headline_metrics(self):
        # Display chevron up for positive numbers or 0, chevron down for negative.
        arrow = self.get_element_by_title('C5 Arrow')
        val = self.get_cell('C5')
        if val >= 0:
            xfrm = self.xpath('.//a:xfrm', arrow)[0]
            xfrm.set('flipV', '1')

    #Fill AVERAGE WEEKS IN RECEIVABLE (FIRM) OK
    def fill_firm_chart(self):
        chart_group = self.E('separate-receivables-firm')

        chart_min_e = self.get_element_by_title('firm-chart-min')
        chart_max_e = self.get_element_by_title('firm-chart-max')

        x0, y0, x1 = self.get_element_coords(chart_max_e)[:3]
        y1 = self.get_element_coords(chart_min_e)[3]

        w = x1-x0 # chart width
        h = y1-y0 # chart height
        step = w/13. # point step

        max_v = 10 # Max heigth
        min_v = 5 # Min height

        def add_max_labels(values, color):
            # If above Max of 10: Show at level of 10, display text above point with actual value in 9 pt font
            capped_values = []

            # adding labels above points 
            for i, v in enumerate(values):
                if v is not None and v > max_v:
                    label = self.clone_template('firm-chart-label')
                    x = x0+(i+.5)*step-self.get_element_sizes(label)[0]/2 # label x position
                    self.set_element_text(label, self.format_float(v, prec=1)) # label text
                    self.set_element_pos(label, x, y0-self.get_element_sizes(label)[1]) # set label to the chart
                    self.set_element_text_color(label, color) # set label color
                    chart_group.append(label)
                    v = max_v
                capped_values.append(v)

            return capped_values

        def add_min_labels(values, color):
            # If below Min of 5: Show at level of 5, display text below point with actual value in 9 pt font
            capped_values = []

            # adding labels above points 
            for i, v in enumerate(values):
                if v is not None and v < min_v:
                    label = self.clone_template('firm-chart-label')
                    x = x0+(i+.5)*step-self.get_element_sizes(label)[0]/2 # label x position
                    self.set_element_text(label, self.format_float(v, prec=1)) # label text
                    # set label to the chart
                    self.set_element_pos(
                        label,
                        x-self.get_element_sizes(label)[0],
                        y1-self.get_element_sizes(label)[1],
                    ) 
                    self.set_element_text_color(label, color) # set label color
                    chart_group.append(label)
                    v = min_v
                capped_values.append(v)

            return capped_values

        chart = self.get_chart_by_title('receivables-firm-chart')
        chart.fill_data(
                [
                    [v for v in [self.get_cell(cell) for cell in C('B%s' % row, 'N%s' % row)]]
                    for row in [10, 11, 12]
                ],
                conv=lambda v: min(max(v, min_v), max_v) if v != 'N/A' else ''
        )
        chart.write()

        for i, row in enumerate([10, 11, 12]):
            color_e = chart.xpath('//a:solidFill/*')[i] # chart line color
            values = [self.get_cell(cell) for cell in C('B%s' % row, 'N%s' % row)] # chart line value
            values = [None if v == 'N/A' else v for v in values] # If value is N/A, do not display line
            values = add_min_labels(add_max_labels(values, color_e), color_e)

    #Fill AVERAGE WEEKS IN RECEIVABLES (BUSINESSES) OK
    def fill_businesses_chart(self):
        # Y-Axis Scale Default to Min=0, Max = 10 
        max_v = 10
        min_v = 0
        y0 = self.get_element_coords(self.get_element_by_title('businesses-chart-max'))[1]
        y1 = self.get_element_coords(self.get_element_by_title('businesses-chart-min'))[1]
        h = y1-y0

        for col in 'BCDE':
            for r in [18, 19, 20]:
                name = '%s%s' % (col, r) # excel cell name
                val_e = self.get_element_by_title(name) # template rectangle value
                rect_e = self.get_element_by_title(name+' Rectangle') # template rectangle
                ln_e = self.xpath('.//a:ln', rect_e)[0] # template rectangle line 
                line_w = 0 if len(self.xpath('.//a:noFill', ln_e, error=False)) else float(ln_e.get('w')) # bottom line position

                # rectangle height: Min height at min_v and Max height at max_v.
                val = self.get_cell(name)
                hn = (max(min(val, max_v), min_v)-min_v)/(max_v-min_v)*h
                y = y1-hn

                self.set_element_pos(val_e, None, y-self.get_element_sizes(val_e)[1]) # set template rectangle value

                self.set_element_pos(rect_e, None, y+line_w/2) # set rectangle position
                self.set_element_size(rect_e, None, hn-line_w) # set rectangle height

            # Set solid line at every third column: Min height at min_v and Max height at max_v.
            e = self.get_element_by_title('%s21' % col)
            hn = (max(min(self.get_cell('%s21' % col), max_v), min_v)-min_v)/(max_v-min_v)*h # solid line height
            y = y1-hn 
            self.set_element_pos(e, None, y-self.get_element_sizes(e)[1]/2)

    #Fill KEY LIQUIDITY METRICS YoY GROWTH OK
    def fill_liquidity_metrics(self):
        chart_group = self.E('separate-key-metrics')
        x0 = self.get_element_coords(self.get_element_by_title('liquidity-metrics-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('liquidity-metrics-max'))[0]
        w = x1-x0 # width between -10% and 25%

        minp = -.10 # Min point position
        maxp = .25 # Max point position
        step = .005 # point step

        for cell in C('B27', 'I27'):
            v = self.get_cell(cell)
            if v < minp:
                #  If Less than -10% Display at -10.5% alignment 
                pos_v = minp-step
            elif v > maxp:
                #  If Greater than 25% Display at 25.5% alignement
                pos_v = maxp+step
            else:
                pos_v = v
            # get green dot
            tick_e = self.get_element_by_title(cell) # green dot 
            tick_w = self.get_element_sizes(tick_e)[0] # green dot width
            x = (pos_v-minp)/(maxp-minp)*w+x0 # green dot position

            # set label with actuale YoY Growth (%) above green dot, if green dot's x is less than -10% or above 25%
            if v < minp or v > maxp:
                label_e = self.clone_template('liquidity-metrics-label')
                label_w, label_h = self.get_element_sizes(label_e)
                self.set_element_text(label_e, self.format_percent(v, prec=1)+'%')
                self.set_element_pos(label_e, x-label_w/2, self.get_element_coords(tick_e)[1]-label_h)
                chart_group.append(label_e)

            self.set_element_pos(tick_e, x-tick_w/2, None)

    #Fill ESTIMATED USES OF CASH NEXT PERIOD
    def fill_estimated_uses_chart(self):
        chart_group = self.E('separate-cash-next-period')
        non_period_total = self.get_cell('C39') # template NON PERIOD DRIVEN CASH NEEDS
        period_total = self.get_cell('C40') # template PERIOD DRIVEN CASH NEEDS
        total = self.get_cell('C41') # template TOTAL

        values = []
        for r in xrange(31, 39):
            row_values = [self.get_cell('%s%s' % (c, r)) for c in 'ABCD'] # excel values
            if row_values[-1]:
                values.append(row_values)
        for v in values:
            v[1] = v[1] == 'Period Driven'
        non_period_values = [v for v in values if not v[1]] # get all non period driven cash needs
        non_period_values.append(['none', None, period_total])
        period_values = [('none', None, non_period_total)] + [v for v in values if v[1]]  # get all period driven cash needs
        angle = 450-non_period_total*180./total if period_total else 0 # calculate pie chart radius

        colors = [
            a.get('val')
            for a in self.get_chart_by_title('estimated-uses-chart-non-period').xpath('//c:dPt//a:srgbClr')[::-1]
        ]

        def fill_chart(chart_name, values, non_data_row):
            # if chart values are empty (no No-Period Driven expenses etc) than remove chart and return
            if len(values) == 1:
                remove(self.get_element_by_title(chart_name))
                return

            chart = self.get_chart_by_title(chart_name) # get pie chart
            label_angle = math.radians(angle)

            # delete all numbers on pie chart
            numCache = chart.xpath('//c:numCache')[0]
            ptCount = chart.xpath('.//c:ptCount', numCache)[0]
            ptCount.set('val', str(len(values)))
            for e in chart.xpath('.//c:pt', numCache):
                remove(e)

            # delete all strings on pie chart
            strCache = chart.xpath('//c:cat/c:strRef/c:strCache')[0]
            ptCount = chart.xpath('.//c:ptCount', strCache)[0]
            ptCount.set('val', str(len(values)))
            for e in chart.xpath('.//c:pt', strCache):
                remove(e)

            ser = chart.xpath('//c:ser')[0]
            for e in chart.xpath('.//c:dPt', ser):
                remove(e)
            dLbls = chart.xpath('.//c:dLbls', ser)[0]
            for e in chart.xpath('c:dLbl', dLbls):
                remove(e)
            chart.xpath('//c:dLbls/c:showVal')[0].set('val', '1')
            chart.xpath('//c:dLbls/c:showLeaderLines')[0].set('val', '0')
            chart.xpath('//c:firstSliceAng')[0].set('val', self.format_float(angle, prec=0))
            txPr = chart.xpath('//c:dLbls/c:txPr')[0]

            coords = self.get_plot_area_coords_from_chart(chart_name)
            r = max((coords[2]-coords[0]), (coords[3]-coords[1]))*.5
            c = ((coords[2]+coords[0])*.5, (coords[3]+coords[1])*.5)

            for i, row in enumerate(values):
                pt = etree.SubElement(strCache, '{%s}pt' % chart.NS['c'])
                pt.set('idx', str(i))
                v = etree.SubElement(pt, '{%s}v' % chart.NS['c'])
                v.text = row[0]

                pt = etree.SubElement(numCache, '{%s}pt' % chart.NS['c'])
                pt.set('idx', str(i))
                v = etree.SubElement(pt, '{%s}v' % chart.NS['c'])
                v.text = self.format_float(row[2]/10.**6, prec=0)

                dPt = etree.SubElement(ser, '{%s}dPt' % chart.NS['c'])
                etree.SubElement(dPt, '{%s}idx' % chart.NS['c']).set('val', str(i))
                spPr = etree.SubElement(dPt, '{%s}spPr' % chart.NS['c'])

                if row is non_data_row:
                    etree.SubElement(spPr, '{%s}noFill' % chart.NS['a'])
                else:
                    solidFill = etree.SubElement(spPr, '{%s}solidFill' % chart.NS['a'])
                    etree.SubElement(solidFill, '{%s}srgbClr' % chart.NS['a']).set('val', colors.pop())

                dLbl = etree.Element('{%s}dLbl' % chart.NS['c'])
                etree.SubElement(dLbl, '{%s}idx' % chart.NS['c']).set('val', str(i))
                etree.SubElement(dLbl, '{%s}showVal' % chart.NS['c']).set('val', '0' if row is non_data_row else '1')
                etree.SubElement(dLbl, '{%s}dLblPos' % chart.NS['c']).set('val', 'inEnd')
                if float(row[2])/total <= .05:
                    txPr_clone = copy.deepcopy(txPr)
                    self.xpath('.//a:defRPr', txPr_clone)[0].set('sz', '1200')
                    dLbl.append(txPr_clone)
                elif row[2]/10.**6 > 9999:
                    etree.SubElement(dLbl, '{%s}layout' % chart.NS['c'])
                    rich_e = etree.SubElement(
                        etree.SubElement(
                            dLbl,
                            '{%s}tx' % chart.NS['c']
                        ),
                        '{%s}rich' % chart.NS['c']
                    )
                    etree.SubElement(
                        rich_e,
                        '{%s}bodyPr' % chart.NS['a'],
                    )
                    etree.SubElement(
                        rich_e,
                        '{%s}lstStyle' % chart.NS['a'],
                    )
                    etree.SubElement(
                        etree.SubElement(
                            etree.SubElement(
                                rich_e,
                                '{%s}p' % chart.NS['a'],
                            ),
                            '{%s}r' % chart.NS['a'],
                        ),
                        '{%s}t' % chart.NS['a']
                    ).text = '$9,999'
                dLbls.insert(i, dLbl)
                element_names = [
                    'showLegendKey',
                    'showCatName',
                    'showSerName',
                    'showPercent',
                    'showBubbleSize',
                    'showLeaderLines',
                ]
                for e_name in element_names:
                    etree.SubElement(dLbl, '{%s}%s' % (chart.NS['c'], e_name)).set('val', '0')

                def get_shift(e, angle):
                    angle -= int(angle)/360*360
                    w, h = self.get_element_sizes(e)
                    pPr = self.xpath('.//a:pPr', e)[0]
                    bodyPr = self.xpath('.//a:bodyPr', e)[0]

                    if 45 <= angle < 135:
                        x = 0
                        pPr.set('algn', 'l')
                    elif 225 <= angle < 315:
                        x = w
                        pPr.set('algn', 'r')
                    else:
                        x = w/2

                    if 45 <= angle < 135 or 225 <= angle < 315:
                        y = h/2
                    elif 135 <= angle < 225:
                        y = 0
                        bodyPr.set('anchor', 'b')
                    else:
                        y = h
                        bodyPr.set('anchor', 't')

                    return (x, y+h/2)  # should be (x, y)

                angle_n = row[2]/float(total)*2*math.pi
                if row is not non_data_row:
                    point_angle = label_angle+angle_n*.5

                    p = (c[0]+r*math.sin(point_angle), c[1]-r*math.cos(point_angle))
                    label = self.clone_template('estimated-uses-chart-label')
                    shift_x, shift_y = get_shift(label, math.degrees(point_angle))
                    self.set_element_text(label, row[0])
                    self.set_element_pos(label, p[0]-shift_x, p[1]-shift_y)
                    chart_group.append(label)
                label_angle += angle_n

            chart.write()

        fill_chart('estimated-uses-chart-non-period', non_period_values, non_period_values[-1])
        fill_chart('estimated-uses-chart-period', period_values, period_values[0])

    def fill_values(self):
        super(Dashboard2, self).fill_values()

        self.fill_headline_metrics()
        self.fill_firm_chart()
        self.fill_businesses_chart()
        self.fill_liquidity_metrics()
        self.fill_estimated_uses_chart()


if __name__ == '__main__':
    CMDHandler(
        Dashboard2,
    )
