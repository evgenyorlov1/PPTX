#!/usr/bin/env python

from __future__ import unicode_literals

import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, CMDHandler


class Dashboard3(PPTXGenerator):

    sheet_name = 'Client Dashboard'
    template_shape_names = (
        'rating-point-dropped',
        'rating-point-new',
        'rating-point-up',
        'rating-point-down',
        'rating-point-same',
        'rating-line-up',
        'rating-line-down',
        'rating-line-same',
    )
    separate_charts = {
        'separate-headline-metrics': '3-Headline-Metrics.pptx',
        'separate-top-clients-by-revenue': '3-Top-Clients-by-Revenue.pptx',
        'separate-top-clients-by-gcm': '3-Top-Clients-by-GCM.pptx',
    }

    #Get values from excel sheet to fill template
    def get_simple_fillers(self):
        # Only display %’s and $ amounts for each portion for Top 15 R’s
        return {
            lambda v: self.format_percent(v) + '%': [C('B4', 'E4')],
            lambda v: self.format_percent(v, prec=0) + '%': [
                C('F25'), C('H25', 'I25'),
                C('F46'), C('H46', 'I46'),
                C('D68', 'E88'),
            ],
            lambda v: '$'+self.format_float(float(v)/10**6, prec=0, maximum=999): [ # Max amounts for Revenue of individual Clients = $999  (if higher display $999)  
                C('D9', 'D23'),
                C('D30', 'D44'),
                C('D50', 'D54'),
                C('D59', 'D63'),
                C('B68', 'C87'),
            ],
            lambda v: '$'+self.format_float(float(v)/10**6, prec=0, maximum=9999): [ # Max Amounts for Revenue of Totals = $9999 (if higher display $9999)
                C('D24'), C('F24'), C('H24', 'I24'),
                C('D45'), C('F45'), C('H45', 'I45'),
                C('D55'),
                C('D64'),
                C('B88', 'C88'),
            ],
            lambda v: v.strip()[:27]: [ # Limit client names to 27 chars including spaces
                C('A9', 'A23'),
                C('A30', 'A44'),
                C('A50', 'A54'),
                C('A59', 'A63'),
                C('A68', 'A87'),
            ],
        }

    def fill_values(self):
        super(Dashboard3, self).fill_values()

        self.add_position_graphics(C('A9', 'A23'), C('A30', 'A44')) # set left chart
        self.add_position_graphics(C('A50', 'A54'), C('A59', 'A63')) # set bottom chart

        # excel cells
        rows = [
            (9, 23),
            (25, None),
            (30, 44),
            (46, None),
            (50, 54),
            (55, None),
            (59, 63),
            (64, None),
        ]

        # chart max value for Axis
        chart_max = max(
            sum(self.get_cell('%s%s' % (c, row)) for c in 'FGHI')
            for row0, row1 in rows if row1
            for row in xrange(row0, row1+1)
        )

        for chart, (row0, row1), in enumerate(rows, 1):
            chart = self.get_chart(self.get_chart_path(chart)) # get template chart

            # fill chart,  Only display portion if higher than 1.5% 
            chart.fill_data(
                [
                    [
                        self.get_cell(name)
                        for name in reversed(C('%s%s' % (c, row0), '%s%s' % (c, row1) if row1 else None))
                    ]
                    for c in 'FGHI'
                ],
                conv=lambda val: 0 if 0 <= val <= 0.015 else val
            )
            chart.set_axis_max(self.format_float(chart_max)) # All other  bar sizes are based off the largest bar 
            chart.write()

    def add_position_graphics(self, cells1, cells2):
        chart_group = self.E('separate-top-clients-by-radd_circleevenue')

        positions0 = {self.get_cell(c): (pos, c) for pos, c in enumerate(cells1)}
        positions1 = {self.get_cell(c): (pos, c) for pos, c in enumerate(cells2)}

        # In 2015 this progresses from right to left 
        def get_right_middle(name):
            coords = self.get_shape_coords(name)
            return coords[2], (coords[1]+coords[3])/2

        # In 2016 it progresses from left to right 
        def get_left_middle(name):
            coords = self.get_shape_coords(name)
            return coords[0], (coords[1]+coords[3])/2

        # If N/A on current list: display a dark green dot, no line 
        for c in positions0.viewkeys()-positions1.viewkeys():
            self.add_circle(*(get_right_middle(positions0[c][1])+('rating-point-dropped',)), shapes=chart_group)

        # if N/A on Prior list: display a red dot, no line 
        for c in positions1.viewkeys()-positions0.viewkeys():
            self.add_circle(*(get_left_middle(positions1[c][1])+('rating-point-new',)), shapes=chart_group)

        #
        for c in positions0.viewkeys() & positions1.viewkeys():
            lt = positions0[c]
            rt = positions1[c]
            # line types: Up, Down, Horizontal
            template_suffix = {
                -1: '-down', # Change to Current for a client is negative: display an orange line connecting the two occurences of that client name between Prior and current year
                0: '-same', # Zero: display a flat, light green line connecting the two names 
                1: '-up', #Positive: display a dark blue line connecting the two occurences of the name between prior and current
            }[cmp(lt[0], rt[0])]

            # add circles for line begining and ending
            self.add_circle(*(get_right_middle(lt[1])+('rating-point'+template_suffix,)), shapes=chart_group)
            self.add_circle(*(get_left_middle(rt[1])+('rating-point'+template_suffix,)), shapes=chart_group)

            # drow line If the Change to Current 
            self.add_line(
                *(get_right_middle(lt[1])+get_left_middle(rt[1])+('rating-line'+template_suffix,)),
                shapes=chart_group
            )

if __name__ == '__main__':
    CMDHandler(
        Dashboard3,
    )
