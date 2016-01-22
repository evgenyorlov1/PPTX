#!/usr/bin/env python

from __future__ import division, unicode_literals

import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, alpha_range, CMDHandler


VARIANCE_NEG_COLOR = '938B95' # Grey hex color
VARIANCE_POS_COLOR = '0A9FDA' # Blue hex color


class Dashboard12(PPTXGenerator):

    sheet_name = 'Enabling Areas Dashboard'
    template_shape_names = (
        'enabling-areas-yoy-down',
        'enabling-areas-yoy-up',
        'total-headcount-label',
        'total-headcount-bubble',
    )
    separate_charts = {
        'separate-headline-metrics': '12-Headline-Metrics.pptx',
        'separate-enabling-areas': '12-Enabling-Areas-Parent-Costs.pptx',
        'separate-cost-breakdown': '12-Cost-Breakdown-For-Parent.pptx',
        'separate-total-ea': '12-Total-EA-and-Parent-Headcount.pptx',
        'separate-ea-and-parent-cost': '12-EA-and-Parent-Cost.pptx',
    }

    def big_money(self, v):
        # MAX amount for any headcount figure is 999,999  
        mlns = round(float(v)/10**6, 0)
        if mlns > 999:
            return '$%sB' % self.format_float(mlns/10**3, prec=2, strip=False)
        else:
            return '$%sM' % self.format_float(mlns, prec=0)

    def get_simple_fillers(self):

        def abs_big_money(v):
            return self.big_money(abs(v))

        return {
            lambda v: abs_big_money(min(v, 99.9*10**9)): [
                C('B3'),
            ],
            lambda v: self.with_comma(min(v, 99999)): [
                C('D3'),
            ],
            lambda v: self.format_percent(min(v, .999), prec=1, strip=False)+'%': [
                C('F3'),
                C('B4'), C('D4'),
            ],
            lambda v: self.format_float(abs(v), prec=0)+'BPS': [
                C('F4'),
                C('F5'),
            ],
            lambda v: ('-' if v < 0 else '+')+'$'+self.format_float(abs(v)/10**6, prec=0)+'M': [
                C('B5'),
            ],
            lambda v: self.format_float(v, prec=0, include_sign=True): [
                C('D5'),
            ],
            lambda v: '$'+self.format_float(min(v, 999*1000)/1000, prec=1)+'K': [
                C('H3'),
            ],
            lambda v: '$'+self.format_float(abs(v)/1000, prec=1)+'K': [
                C('H4'),
            ],
            lambda v: self.format_float(v/1000, prec=1, include_sign=True)+'K': [
                C('H5'),
            ],
            lambda v: '$'+self.format_float(v/10.**6, prec=0): [
                C('B10', 'J10'),
            ],
            lambda v: ('(%s)' if v < 0 else '%s') % ('$'+self.format_float(abs(v)/10.**6, prec=0)): [
                C('B20', 'G20'),
            ],
            self.big_money: [
                C('B8'),
            ],
            lambda v: self.with_comma(min(v, 999999)): [
                C('B31', 'E31'),
                C('B33', 'D33'),
                C('B34', 'D34'),
            ],
            lambda v: self.format_percent(v, prec=0)+'%': [
            ],
        }

    def get_lines_fillers(self):
        # set labels
        return [
            (9, [
                C('B9', 'J9'),
                C('B19', 'G19'),
            ]),
        ]

    def fill_headline_metrics(self):
        pass

    # Enabling Areas & Parent Costs 
    def fill_enabling_areas(self):
        separate_chart_group = self.E('separate-enabling-areas')

        y1 = self.get_element_coords(self.E('enabling-areas-min'))[3]
        max_h = self.get_element_sizes(self.E('B11'))[1] # Largest value will be Max height

        for col in alpha_range('B', 'J'): 
            bar_e = self.E('%s11' % col) # section rectangle Bar Height (% of Max) 
            bar_coords = self.get_element_coords(bar_e)
            bar_hs = [] # sections height

            bar_h = max_h*self.get_cell('%s11' % col) # normalized height
            bar_hs.append(bar_h)
            if bar_h < 0:
                # If negative value, display no bar, and display value just above x axis 
                remove(bar_e)
            else:
                # set section position
                self.set_element_pos(bar_e, None, y1-bar_h)
                self.set_element_size(bar_e, None, bar_h)

            for name in ['%s13' % col, '%s15' % col]:
                # lines height
                bar_h = self.get_cell(name)*max_h # section height
                bar_hs.append(bar_h)
                e = self.E(name)
                if bar_h < 0:
                    # # If negative value, display no bar, and display value just above x axis 
                    remove(e)
                else:
                    self.set_element_pos(e, None, y1-bar_h)

            yoy = self.get_cell('%s16' % col) # YoY Cost Growth
            yoy_e = self.clone_template('enabling-areas-yoy-down' if yoy < 0 else 'enabling-areas-yoy-up')
            yoy_e_w, yoy_e_h = self.get_element_sizes(yoy_e) # label width, height
            # set YOY label
            self.set_element_pos(
                yoy_e,
                (bar_coords[0]+bar_coords[2]-yoy_e_w)/2,
                self.get_element_coords(bar_e)[1]-yoy_e_h
            )
            self.set_element_text(yoy_e, self.format_percent(abs(yoy), prec=0)+'%')
            separate_chart_group.append(yoy_e)

    # Cost Breakdown For Parent
    def fill_cost_breadown(self):
        y1 = self.get_element_coords(self.E('cost-breakdown-min'))[3]
        max_h = self.get_element_sizes(self.E('B22'))[1] # Largest value will be Max height

        for col in alpha_range('B', 'G'):
            bar_e = self.E('%s22' % col) # section rectangle Bar Height (% of Max) 
            bar_coords = self.get_element_coords(bar_e)
            bar_hs = [] # sections height

            def get_h(p):
                # Limit the height of any negative values to -1/4 of MAX 
                h = max_h*max(p, -.25)
                bar_hs.append(h)
                return h

            bar_h = get_h(self.get_cell('%s22' % col))
            if bar_h < 0: # Limit the height of any negative values to -1/4 of MAX *
                self.set_element_text_color(self.E('%s20' % col), 'FFFFFF')
            # set Current Amt rectangle
            self.set_element_pos(bar_e, None, y1-max(bar_h, 0))
            self.set_element_size(bar_e, None, abs(bar_h))

            for name in ['%s24' % col, '%s26' % col]:
                e = self.E(name)
                self.set_element_pos(e, None, y1-get_h(self.get_cell(name)))

            yoy = self.get_cell('%s21' % col) # YoY Growth(%)
            yoy_e = self.clone_template('enabling-areas-yoy-down' if yoy < 0 else 'enabling-areas-yoy-up')
            yoy_e_w, yoy_e_h = self.get_element_sizes(yoy_e) # YOY width and height
            # set YOY label values
            self.set_element_pos(
                yoy_e,
                (bar_coords[0]+bar_coords[2]-yoy_e_w)/2,
                self.get_element_coords(bar_e)[1]-yoy_e_h
             )
            self.set_element_text(yoy_e, self.format_percent(abs(yoy), prec=0)+'%')
            self.add_shape(yoy_e)

            # section label below x Axis
            label_e = self.E('%s19' % col)
            self.set_element_pos(
                label_e,
                None,
                max(self.get_element_coords(self.E('%s%s' % (col, x)))[3] for x in [20, 22, 24, 26]),
            )

        self.set_element_text(self.E('I20'), self.big_money(self.get_cell('B10'))) # Total

    # Total EA & Parent Headcount
    def fill_parent_headcount(self):
        max_val = self.get_cell('B29') # Max of Scale
        labels = [('B29*%s/5' % i, max_val*i/5.) for i in range(1, 5)]+[('B29', max_val)] # numerate x Axis
        for name, value in labels:
            self.set_text(name, self.format_float(value/10**3, prec=0)+'K') # format x Axis numeration (50K, 30K etx)

        # set values under arrows(%)
        for cell in C('B32', 'E32')+C('B35', 'D35'):
            self.set_element_text(self.E(cell), self.format_percent(self.get_cell(cell), prec=0)+'%') 

        x0 = self.get_element_coords(self.E('parent-headount-min'))[0]
        x1 = self.get_element_coords(self.E('parent-headount-max'))[0]
        w = x1-x0 # chart width

        # set EA & Parent Headcount and Total Firm Headcount 
        # and Prior YTD and Prior Firm Headcount
        for cell in C('B31', 'D31')+C('B33', 'D33')+C('T48', 'V48')+C('B34', 'D34'):
            e = self.E('%s-rect' % cell)
            wpn = self.get_cell(cell)/max_val # section width normalized
            self.set_element_size(e, wpn*w, None)

        for titles, yoy_title in zip(zip(C('B33', 'D33'), C('B34', 'D34')), C('B35', 'D35')):
            bar_e = self.E('%s-rect' % titles[0])
            bar_coords = self.get_element_coords(bar_e) # section coords
            wpn = (bar_coords[2]-bar_coords[0])/w # normalized section width
            if wpn < .35:
                # If dotted line section is less than 35% of scale max 
                x = bar_coords[2] # Display Value of total headcount at right edge of dotted line box 
                yoy_x = bar_coords[2]+self.get_element_sizes(self.E(titles[0]))[0] # Display YoY growth to the right of total headcount value 
            else:
                # If dotted line section (Total Headcount - EA/Parent Headcount) is greater than 35%
                x = (bar_coords[0]+bar_coords[2])/2 # Display Total Headcount value centered within dotted line section for current year
                yoy_x = bar_coords[2] # Display YoY growth rate at the right edge of the bar 

            self.set_element_pos(self.E(yoy_title), yoy_x, None)

            # section titles
            for title in titles:
                # If dotted line section is less than 35% of scale max 
                # Display Value of total headcount at right edge of dotted line box 
                # Display YoY growth to the right of total headcount value
                e = self.E(title)
                self.set_element_pos(e, x-(0 if wpn < .35 else self.get_element_sizes(e)[0])/2, None)

    # EA Total Headcount & Cost
    def fill_total_headcount(self):
        separate_chart_group = self.E('separate-ea-and-parent-cost')

        max_size = self.get_element_sizes(self.template_shapes['total-headcount-bubble'])[0] # max chart size
        font_size = self.get_element_font_size(self.template_shapes['total-headcount-bubble'])
        x_axis_coords = self.get_element_coords(self.E('total-headcount-x-axis')) # x Axis
        y_axis_coords = self.get_element_coords(self.E('total-headcount-y-axis')) # y Axis

        x0 = x_axis_coords[0]
        w = x_axis_coords[2]-x0 # chart width

        y1 = y_axis_coords[3]
        h = y1-y_axis_coords[1] # chart height

        # x Axis scale Min -25% and Max 25%
        min_xp = -.25
        max_xp = .25

        # y Axis scale Min -15% and Max 15%
        min_yp = -.15
        max_yp = .15

        for col in alpha_range('B', 'H'):
            size_p = self.get_cell('%s43' % col)
            if size_p < .25:
                continue

            size = size_p*max_size # label size
            bubble_e = self.clone_template('total-headcount-bubble') # label
            self.set_element_size(bubble_e, size, size)
            # Limit the positions of the bubbles to a max of 26% and a min of -26% on the x axis
            x_val = max(min(self.get_cell('%s42' % col), .26), -.26)
            # Limit the positions of the bubbles to max of 20% and a min of -20% on the y axis.
            y_val = max(min(self.get_cell('%s40' % col), .2), -.2)
            self.set_element_pos(
                bubble_e,
                x0 + (x_val-min_xp)/(max_xp-min_xp)*w-size/2,
                y1 - (y_val-min_yp)/(max_yp-min_yp)*h-size/2,
            )
            self.set_element_text(bubble_e, self.get_cell('%s38' % col)) # label text(IT, SINT etc)
            self.set_element_font_size(bubble_e, size_p*font_size) # set bubble  size
            separate_chart_group.append(bubble_e)

            # Total Headcount and Total Cost labels
            label_e = self.clone_template('total-headcount-label')
            self.set_element_text_lines(
                label_e,
                [
                    'Cost: $%sM' % self.format_float(self.get_cell('%s41' % col)/10**6, prec=0),
                    'TOTAL HC: %s' % self.format_float(self.get_cell('%s39' % col), prec=0),
                ]
            )
            self.set_element_pos(
                label_e,
                self.get_element_coords(bubble_e)[0]-self.get_element_sizes(label_e)[0],
                self.get_element_coords(bubble_e)[1]+size/2-self.get_element_sizes(label_e)[1]/2,
            )
            separate_chart_group.append(label_e)

    def fill_values(self):
        super(Dashboard12, self).fill_values()

        self.fill_headline_metrics()
        self.fill_enabling_areas()
        self.fill_cost_breadown()
        self.fill_parent_headcount()
        self.fill_total_headcount()


if __name__ == '__main__':
    CMDHandler(
        Dashboard12,
    )
