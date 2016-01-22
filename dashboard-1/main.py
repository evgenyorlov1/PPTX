#!/usr/bin/env python

from __future__ import division, unicode_literals

import math
import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, alpha_range, CMDHandler


class Dashboard1(PPTXGenerator):

    sheet_name = 'Service Area Dashboard'
    template_shape_names = (
        'template-arrow',
    )
    separate_charts = {
        'separate-headline-metrics': '1-Service-Area-Headline-Metrics.pptx',
        'separate-revenue-eba': '1-Service-Area-Rev-Eba.pptx',
    }

    #Get values from excel sheet to fill template OKs
    def get_simple_fillers(self):
        #Get values from excel sheet to fill template.
        return {
            lambda v: v: [
                C('F2'), C('H2'), C('K2'),
                C('B6', 'M6'),
            ],
            lambda v: self.format_percent(abs(v), prec=0)+'%': [
                C('C4'), C('F4'), C('H4'), C('K4'),
            ],
            lambda v: self.format_percent(abs(v), prec=0)+'%': [
                C('B9', 'M9'),
            ],
            lambda v: self.format_percent(v, prec=0)+'%': [
                C('B12', 'M12'),
            ],
            lambda v: '$'+self.format_float(v/10.**6, prec=0): [
                C('B7', 'M7'),
                C('B10', 'M10'),
            ]
        }

    def fill_headline_metrics(self):
        #If growth rate is negative: have chevron pointing down, else up.
        for c in C('C4')+C('F4')+C('H4')+C('K4'):
            val = self.get_cell(c)
            remove(self.get_element_by_title('%s-arrow-%s' % (c, 'up' if val < 0 else 'down')))

        #In 'Advisory' write 'Analytics' instead of 'Risk Analytics'.
        C2 = self.get_cell('C2')
        if C2 == 'Risk Analytics':
            C2 = 'Analytics'
        self.set_text('C2', C2)

    def fill_chart(self):
        cols = alpha_range('B', 'M') # excel cells in range.

        x0, y0, x1, y1 = self.get_element_coords(self.get_element_by_title('chart-box'))
        w = x1-x0 # chart box width
        h = y1-y0 # chart box height

        self.set_element_pos(self.get_element_by_title('eba'), None, y0+h*.75)

        x = x0
        for c in cols:
            wpn = self.get_cell('%s8' % c)
            wn = wpn*w
            hpn = min(max(self.get_cell('%s11' % c), .08), .83) # solid box height is between 8% and 83%
            hn = hpn*h

            rotated = wpn <= .04

            rect = self.get_element_by_title('%s11' % c) # solid box
            eba_label = self.get_element_by_title('%s10' % c) # Earnings
            growth_e = self.get_element_by_title('%s9' % c) # Revenue Growth 
            revenue_label = self.get_element_by_title('%s7' % c) # Revenue
            title = self.get_element_by_title('%s6' % c) # column title
            arrow = self.clone_template('template-arrow') # get arrow chevron

            self.set_element_size(rect, wn, hn) # set solid box size
            self.set_element_pos(rect, x, y0+h-hn) # set solid box position

            # Rules for displaying EBA ($) figure, solid box:
            if hpn >= .32:
                # If EBA Margin is between 83%-32%, display EBA ($) in white with top of text at 32% height
                self.set_element_pos(eba_label, None, y0+h-.32*h)
            elif hpn >= .27:
                # If EBA Margin is 32%-27% Display EBA ($) in the color of the function (grey, dark blue, light blue, or green) with top of text at 34.5% height
                self.set_element_text_color(eba_label, self.get_element_fill_color(rect))
                self.set_element_pos(eba_label, None, y0+h-.345*h)
            else:
                # If EBA Margin in 27%-8%, display EBA ($) in the color of the function (grey, dark blue, light blue, or green) with top of text at 32% height
                self.set_element_text_color(eba_label, self.get_element_fill_color(rect))
                self.set_element_pos(eba_label, None, y0+h-.32*h)

            # right-top corner of title box should be at center of section horizontally
            title_w, title_h = self.get_element_sizes(title)
            title_angle = math.radians(self.get_element_rotation(title)/60000)
            # calculating horizonatal shift of right-top corner of box from its center
            title_shift = (title_w*math.cos(title_angle)+title_h*math.sin(title_angle))/2
            self.set_element_pos(title, x+(wn-title_w)/2-title_shift, None)

            # Rules for displaying Revenue Growth (%) chevron:
            if wpn > .015:
                # When width of a section is above 1.5%, the chevron of Revenue Growth(up/down) should appear
                self.set_element_pos(arrow, x+(wn-self.get_element_sizes(arrow)[0])/2, None)
                self.E('separate-revenue-eba').append(arrow)

                if self.get_cell('%s9' % c) < 0:
                    # When Revenue Growth is positive, chevron goes up.
                    self.set_element_flipv(arrow, True)
                    self.set_element_fliph(arrow, True)

            # For Width of columns, except First column.  
            for l in [7, 9, 10, 12]:
                label = self.get_element_by_title('%s%s' % (c, l))
                if wpn <= .015:
                    # If Width is below 1.5%, Do not display any  numbers on the column
                    remove(label)
                else:
                    self.set_element_size(label, wn, None)
                    self.set_element_pos(label, x, None)
                    # If Larger than 3.5%, just do typical horizontal type with 13.6 size font.
                    #If Between 3.5% and 1.5%, turn all text horizontal but keep at same size font
                    self.set_element_text_direction(label, 'vert270' if .015 < wpn < .035 else None)
                    self.set_element_text_alignment(label, 'r' if .015 < wpn < .035 else 'ctr')

            # For Width of the First column.
            if c == cols[0]:
                if rotated:
                    #  If smaller than 4% wide, move labels (Revenue,  EBA, Margin)  outside the left edge of the chart.
                    for e_title in ['eba', 'revenue', 'margin']:
                        e = self.get_element_by_title(e_title)
                        self.set_element_pos(e, x0-self.get_element_sizes(e)[0], None) # to the left edge
                        self.set_element_text_color(e, '000000') # set them black color
                else:
                    # If width larger than 4%:
                    if hpn < .25:
                        # If margin is below 25%, display EBA $ at the normal 25% level and change color to match the box below it. 
                        self.set_element_text_color(
                            self.get_element_by_title('eba'),
                            self.get_element_fill_color(rect),
                        )
                    #  If margin is higher than 80%, change color of the Revenue # Amount to white, move 25% above Height 
                    elif hpn > .8:
                        revenue_e = self.get_element_by_title('revenue')
                        self.set_element_pos(revenue_e, None, y0+.25*h)
                        self.set_element_pos(revenue_label, None, y0+.25*h-self.get_element_sizes(revenue_label)[1])
                        self.set_element_text_color(revenue_label, 'FFFFFF')
                        self.set_element_text_color(revenue_e, 'FFFFFF')

            x += wn

            if c != cols[-1]:
                sep = self.get_element_by_title('%s-sep' % c)
                self.set_element_pos(sep, x, None)


        # Set advisory, audit, consulting and tax in the middle of the column.
        for label_name, (first, last) in [
            ('advisory', 'BE'),
            ('audit', 'FF'),
            ('consulting', 'GI'),
            ('tax', 'JM'),
        ]:
            label = self.get_element_by_title(label_name)
            start = self.get_element_by_title('%s11' % first)
            end = self.get_element_by_title('%s11' % last)
            self.set_element_pos(
                label,
                (self.get_element_coords(end)[2]+self.get_element_coords(start)[0]-self.get_element_sizes(label)[0])/2,
                None,
            )

    def fill_values(self):
        super(Dashboard1, self).fill_values()

        self.fill_headline_metrics()
        self.fill_chart()


if __name__ == '__main__':
    CMDHandler(
        Dashboard1,
    )
