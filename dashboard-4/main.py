#!/usr/bin/env python

from __future__ import unicode_literals

import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, alpha_range, CMDHandler


class Dashboard4(PPTXGenerator):

    sheet_name = 'Industry and Sector Dashboard'
    template_shape_names = (
        'small-arrow-down',
        'small-arrow-up',
        'big-arrow-down',
        'big-arrow-up',
    )
    separate_charts = {
        'separate-headline-metrics': '4-Headline-Metrics.pptx',
        'separate-industry-revenue-mix': '4-Industry-Revenue-Mix.pptx',
    }

    #Get values from excel sheet to fill template
    def get_simple_fillers(self):
        # formating values for Revenue for indastry charts
        def mln_wit_sep(v):
            s = self.format_float(float(v)/10**6, prec=0)
            if len(s) > 3:
                s = s[:-3]+','+s[-3:]
            return '$'+s

        return {
            lambda v: v: [
                C('B3'), C('D3'), C('F3'), C('H3')
            ],
            lambda v: self.format_percent(v, prec=0) + '%': [
                C('B4'), C('D4'), C('F4'), C('H4')
            ],
            lambda v: '$'+self.format_float(float(v)/10**6, prec=1, strip=False)+'M': [
                C('B5'), C('D5'), C('F5'), C('H5')
            ],
            mln_wit_sep: [
                C('B13', 'H13')
            ],
        }

    # Fill REVENUE BY INDUSTRY SECTOR AND BUSINESS
    def fill_small_charts(self):
        # fill small charts at the top, one-by-one
        for l, c in zip('BCDEFGH', ['cip', 'er', 'fed', 'fs', 'lshc', 'ps', 'tmt']):
            chart = self.get_chart_by_title('%s-chart' % c) # get small chart
            chart.fill_data([[self.get_cell(name)] for name in C('%s9' % l, '%s12' % l)]) # fill small chart with values
            chart.write()

            val = self.get_cell('%s14' % l)
            chart_element = self.get_element_by_title('%s-chart' % c)
            arrow = self.clone_template('big-arrow-down' if val < 0 else 'big-arrow-up') #
            # If Revenue Growth is negative, display chevron down, else up
            self.set_element_pos(
                arrow,
                self.get_element_coords(chart_element)[2]-self.get_element_sizes(arrow)[0],
                None
            )
            # set % above the arrow
            self.set_element_text(
                self.xpath('*[*/p:cNvPr[@title="value"]]', arrow)[0],
                self.format_percent(abs(val), prec=0)+'%'
            )
            self.E('separate-industry-revenue-mix').append(arrow)

    def fill_values(self):
        super(Dashboard4, self).fill_values()

        def to_zero(v):
            return 0 if v < 0.015 else v

        self.fill_small_charts()
        # fill bottom chart with values
        self.fill_chart(
            2,
            [
                [self.get_cell(name) for name in C('B%s' % c, 'Z%s' % c)]
                for c in range(19, 23)
            ],
            conv=to_zero,
        )

        # heights based on advisory, audit, consulting and tax
        heights = [
            sum(to_zero(self.get_cell('%s%s' % (c, r))) for r in range(19, 23))
            for c in alpha_range('B', 'Z')
        ]

        # numerate y-Axis
        B17 = self.get_cell('B17') # y-Axis maximum
        # y-Axis points
        for i, c in enumerate(['B17*1/5', 'B17*2/5', 'B17*3/5', 'B17*4/5', 'B17'], 1):
            self.set_text(c, '$'+self.format_float(B17*i/5/10**6, prec=1))

        # chart coords for placing arrows
        chart = self.get_chart_by_title('chart')
        coords = self.get_element_coords(self.get_element_by_title('chart'))
        chart_width, chart_height = self.get_element_sizes(self.get_element_by_title('chart'))
        x0 = coords[0]+chart_width*chart.x
        y0 = coords[3]-chart_height*chart.y
        width = chart_width*chart.w # chart width
        height = chart_height*chart.h # chart height
        n = 25 # num of small arrows\rectangles 
        step = width/n # step between small arrows

        # small arrows on a big chart
        for i, (c, h) in enumerate(zip(C('B23', 'Z23'), heights)):
            val = self.get_cell(c)
            tmpl_name = 'small-arrow-up' if val >= 0 else 'small-arrow-down' # If Revenue Growth is negative, display chevron up, else - down
            tmpl = self.clone_template(tmpl_name)
            coords = self.get_element_coords(tmpl)
            tmpl_w, tmpl_h = self.get_element_sizes(tmpl)
            self.set_element_text(tmpl, self.format_percent(abs(val), prec=0)+'%') # set text under small arrow
            # place arrowsabove rectangles
            self.set_element_pos(
                tmpl,
                x0+step*i+(step-tmpl_w)/2+(tmpl_w/2+step/4)*(h > .95),
                y=y0-height*h+(-tmpl_h*(h <= .95)),
            ) # set small arrow
            self.E('separate-industry-revenue-mix').append(tmpl)


if __name__ == '__main__':
    CMDHandler(
        Dashboard4,
    )
