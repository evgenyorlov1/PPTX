#!/usr/bin/env python

from __future__ import unicode_literals

import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, CMDHandler, alpha_range


VARIANCE_NEG_COLOR = '938B95' # Grey
VARIANCE_POS_COLOR = '0A9FDA' # Blue


class Dashboard7(PPTXGenerator):

    sheet_name = 'Consulting Dashboard'
    template_shape_names = (
        'marg-arrow-up',
        'marg-arrow-down',
        'top-10-arrow',
        'top-10-arrow-max',
        'top-10-arrow-label',
    )
    separate_charts = {
        'separate-headline-metrics': '7-Headline-Metrics.pptx',
        'separate-revenue-and-earnings': '7-Revenue-Earnings-by-Major-Service-Area.pptx',
        'separate-key-metrics': '7-Key-Metrics.pptx',
        'separate-eba-earnings': '7-EBA-Earnings.pptx',
        'separate-rates-and-cs-headcount': '7-Rates-CS-Headcount-by-Geo.pptx',
        'separate-top-10-managed-clients': '7-Top-10-Managed-Clients.pptx',
    }

    def get_simple_fillers(self):
        #If below $999 Million, display as millions. (i.e.$987M)
        #If above  $1000M, display as Billions, with two decimal places (i.e. $1.34B  or $2.00B)
        def big_money(v):
            mlns = round(float(v)/10**6, 0)
            if mlns > 999:
                return '$%sB' % self.format_float(mlns/10**3, prec=2, strip=False)
            else:
                return '$%sM' % self.format_float(mlns, prec=0)

        def abs_big_money(v):
            return big_money(abs(v))

        return {
            lambda v: big_money(min(v, 99.99*10**9)): [
                C('B3'), C('D3'), C('F3'),
            ],
            big_money: [
                C('B4'), C('D4'), C('F4'),
                C('B8'),
            ],
            lambda v: self.format_money(v/10.**6, prec=0): [
                C('B31', 'N31'),
                C('B33'),
            ],
            lambda v: str(int(v)): [
                C('B30', 'N30'),
                C('B38', 'AN38'),
            ],
            lambda v: v: [
                C('A51', 'A60'),
                C('B9', 'D9'),
            ],
            lambda v: self.format_percent(min(v, .9999), prec=1, strip=False)+'%': [
                C('H3'), C('J3'),
                C('B5'), C('D5'), C('F5'),
            ],
            lambda v: '%sBPS' % self.format_float(abs(v), prec=0): [
                C('H4'), C('J4'),
                C('H5'), C('J5'),
            ]
        }

    # REVENUE & EARNINGS BY MAJOR SERVICE AREA
    def fill_blue_rect(self):
        separate_chart_group = self.E('separate-revenue-and-earnings')

        def rotate(e):
            # rotate elements(90 degrees)
            self.set_element_rotation(e, -5400000)
            w, h = self.get_element_sizes(e)
            shift = (h-w)/2
            coords = self.get_element_coords(e)
            x, y = coords[0], coords[1]
            x += shift
            y -= shift
            self.set_element_pos(e, x, y)
            self.set_element_text_alignment(e, 'r')

        coords = self.get_element_coords(self.shapes['big-blue-rect']) # base rectangle coords
        w, h = coords[2]-coords[0], coords[3]-coords[1]
        x0, y0 = coords[0], coords[3]
        # minor rectangles sizes
        sizes = zip(
            'BCD',
            [self.get_cell(c) for c in C('B12', 'D12')],
            [self.get_cell(c) for c in C('B16', 'D16')],
        )
        # If below 3.0% width: Do not show section at all. 
        sizes = [(l, wn, hn) if wn > 0.03 else (l, 0., 0.) for l, wn, hn in sizes]
        norm = sum(wn for l, wn, hn in sizes)

        x = x0
        for l, wpn, hpn in sizes:
            wn = wpn*w/norm

            if wpn > .11:
                # If section is above 11% width (row 12): 
                # Max height of Earnings box (light blue) is 80%
                # Min Height of Earnings box (light blue) is 13.33%
                hpn_min, hpn_max = .1333, .8

                def format_money(v):
                    if v > 9999:
                        return '$%sB' % self.format_float(v/10**3, prec=1)
                    return '$%s' % self.format_float(v, prec=0)
            elif wpn > .065:
                # If section is below 10% width, and above 6.5% width: 
                # Min Height of Earnings box (light blue) is 20.5% 
                # Max height of earnings box (light blue) is 72%
                hpn_min, hpn_max = .205, .72

                def format_money(v):
                    return '$%s' % self.format_float(min(v, 999), prec=0)
            elif wpn > .03:
                # If below 6.5% width, above 3.0% width:
                # Max height of earnings box is 77% 
                # Minimum height of earnings box is 0%
                hpn_min, hpn_max = 0, .77
            else:
                # Do not show section at all. 
                hpn = 0

            hn = min(max(hpn, hpn_min), hpn_max)*h if wpn else 0

            # set rectangle
            r = self.shapes['%s-rect' % l]
            if wpn == 0:
                remove(r)
            else:
                 # set rectangle size and position
                self.set_element_pos(r, x, y0-hn)
                self.set_element_size(r, wn, hn)

            # set section title(A&F, ANLYT, Bussiness Risk, Tech Risk)
            section_title = self.get_element_by_title('%s9' % l)
            if wpn == 0:
                # if width == 0, remove title
                remove(section_title)
            else:
                # set rectangle size and position
                self.set_element_pos(section_title, x, coords[1])
                if wpn <= .125:
                    # If width is below 12.5%, rotate name vertically ​​
                    if wn < self.get_element_sizes(section_title)[1]:
                        self.set_element_text_inset(section_title, 't', 0)
                        self.set_element_text_inset(section_title, 'b', 0)
                        self.set_element_text_vert_alignment(section_title, 'ctr')
                        self.set_element_size(section_title, None, wn)
                    rotate(section_title)

            # set earnings
            earnings = self.get_element_by_title('%s13' % l)
            if wpn <= .065:
                remove(earnings)
            else:
                # if width > 11%: 
                # if height of earnings box is above 30%: display Earnings value($) in white, at 29.15%  
                # if 6.5% < width < 11% :
                # If earnings box is above 38%: display earnings  value ($)in white, at 37.5%
                # If earnings box is below 38%: display earnings  value($) in light blue, at 53.33%  
                earnings_value = self.get_cell('%s13' % l)/10.**6
                t, low, high = (.3, .2915, .3677) if wpn > .11 else (.38, .375, .5333)
                earnings_hp = low if hpn > t else high
                color = 'FFFFFF' if hpn > t else'00B0F0' # set color of ERN label
                # set earnings label position, text and color 
                self.set_element_pos(earnings, x, y0-earnings_hp*h)
                self.set_element_text(earnings, format_money(earnings_value))
                self.set_element_text_color(earnings, color)

                # if section width below 11% rotate it
                if wpn < .11 and round(earnings_value, 0) >= 100:
                    rotate(earnings)

            revenue_rotated = False
            margin_rotated = False

            for c in [10, 14]:
                title = '%s%s' % (l, c) #rectangle values
                val = self.get_cell(title)
                if c == 10:
                    val /= 10.**6 # format REV value
                s = self.get_element_by_title(title)
                if wpn <= .065:
                    remove(s) # remove values if width < 6.5%
                else:
                    # set REV label on template
                    self.set_element_pos(s, x, None)
                    self.set_element_text(s, format_money(val) if c == 10 else self.format_percent(val, prec=0)+'%') # format values
                    if c == 10 and wpn < .125:
                        # If width <12.5%, rotate name vertically, and move Revenue ($) at 70%​​
                        self.set_element_pos(s, None, y0-.7*h)
                    if wpn < .11 and (c == 10 and round(val, 0) >= 100 or val == 1):
                        if c == 14:
                            margin_rotated = True
                        else:
                            revenue_rotated = True
                        rotate(s)

            # set REV arrow
            rev_arrow = self.get_element_by_title('%s11-arrow' % l)
            if wpn <= .065:
                # If below 6.5% width, remove arrow
                remove(rev_arrow)
            else:
                self.set_element_text(rev_arrow, self.format_percent(self.get_cell('%s11' % l), prec=0)+'%')
                rev_w, rev_h = self.get_element_sizes(self.get_element_by_title('%s10' % l))
                shift = rev_w-self.get_element_sizes(rev_arrow)[0]
                self.set_element_pos(rev_arrow, x+shift/2, None)
                # If width < 12.5%, move Revenue Growth(%) at 55%
                if wpn < .125:
                    self.set_element_pos(rev_arrow, None, y0-.55*h)
                if revenue_rotated:
                    rot_shift = (rev_h - rev_w)/2
                    self.mod_element_pos(rev_arrow, rot_shift, -rot_shift)

            # set MARG arrow
            if wpn > .065:
                marg_arrow_val = self.get_cell('%s15' % l) # MARG value
                marg_arrow = self.clone_template('marg-arrow-up' if marg_arrow_val >= 0 else 'marg-arrow-down') # MARG arrow up/down
                self.set_element_text(marg_arrow, self.format_float(min(abs(marg_arrow_val), 999), prec=0)+'BPS') # set MARG value(BPS)
                marg_w, marg_h = self.get_element_sizes(self.get_element_by_title('%s14' % l)) # MARG value(%)
                shift = marg_w-self.get_element_sizes(marg_arrow)[0]
                self.set_element_pos(marg_arrow, x+shift/2, None)
                separate_chart_group.append(marg_arrow)
                if margin_rotated:
                    rot_shift = (marg_h - marg_w)/2
                    self.mod_element_pos(marg_arrow, rot_shift, -rot_shift)

            x += wn

            if l != 'D':
                s = self.get_elements_by_title('%s-sep' % l)[0]
                if wn == 0:
                    remove(s)
                else:
                    self.set_element_pos(s, x, None)

    # RATES & CS HEADCOUNT BY GEOGRAPHY
    def fill_rates_by_geography(self):
        f_func = self.with_comma
        h_func = lambda v: '$'+self.format_float(v, prec=0)

        # numerate left y-Axis
        for l, f in [('F', f_func), ('H', h_func)]:
            v = self.get_cell('%s36' % l)
            values = [v*i/6 for i in range(6, 0, -1)]
            cells = ['%s36' % l] + ['%s36*%s/6' % (l, i) for i in range(5, 0, -1)]

            for v, c in zip(values, cells):
                self.set_text(c, f(v))

        # Fill India Rate, Mexico Rate, Average Rate and US Rate lines
        lines_chart = self.get_chart_by_title('rates-by-geography-chart-lines')
        lines_chart.fill_from_cells(
            [C('B%s' % n, 'AN%s' % n) for n in [42, 43, 44, 45]]
        )
        # Chart right y-axis min/max
        lines_chart.set_axis_min('0')
        lines_chart.set_axis_max(self.format_float(self.get_cell('H36')))
        lines_chart.write()

        area_chart = self.get_chart_by_title('rates-by-geography-chart-area')
        area_chart.fill_from_cells(
            [C('B%s' % n, 'AN%s' % n) for n in [39, 40, 41]]
        )
        # Chart left y-axis min/max
        area_chart.set_axis_min('0')
        area_chart.set_axis_max(self.format_float(self.get_cell('F36')))
        area_chart.write()

    # KEY METRICS
    def fill_key_metrics(self):
        x0 = self.get_element_coords(self.get_element_by_title('key-metrics-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('key-metrics-max'))[0]
        w = x1-x0 # chart width

        def set_tick(name, disable=False):
            tick = self.get_element_by_title(name+'-tick')
            value_e = self.get_element_by_title(name)
            # If both icons are either above 15% or below -10% on the same metric, 
            # only display actual amount for advisory icon
            if disable:
                remove(value_e)

            # tick positioning
            v = self.get_cell(name)
            tick_w = self.get_element_sizes(tick)[0]
            if v < -.10: # If  amount is lower than -10%: Display at -10.5% on scale, display actual amount  (with no decimal point) above the icon
                pos_v = -.105
            elif v > .15: # If amount is higher than 15%: Display at 15.5% on scale, display actual amount (with no decimal point) above the icon
                pos_v = .155
            else:
                pos_v = v # If both icons are either above 15% or below-10% on the same metric, only display actual amount for advisory icon

            x = x0 + (min(max(pos_v, -0.105), .155)+.1)*4*w # tick coord
            self.set_element_pos(tick, x-tick_w/2, None) # set tick on the chart

            value_e_w = self.get_element_sizes(value_e)[0] # tick label

            # if higher than 99%, just display 99%;  if Lower than -99%, just display -99% 
            if -0.99 < v < 0.99:
                text = self.format_percent(v, prec=1)+'%'
            else:
                text = self.format_percent(min(max(v, -0.99), .99), prec=0)+'%'
            self.set_element_text(value_e, text)
            self.set_element_pos(value_e, x-value_e_w/2, None)

        # Place ticks on chart
        for n in range(21, 26):
            b_name = 'B%s' % n # Advisory
            c_name = 'C%s' % n # Firm

            bv = self.get_cell(b_name)
            cv = self.get_cell(c_name)

            set_tick(b_name, disable=-.10 <= bv <= .15) # set advisory tick
            set_tick(c_name, disable=-.10 <= cv <= .15 or all(v < -.10 for v in [bv, cv]) or all(v > .15 for v in [bv, cv])) # set Firm tick

    # ADVISORY EBA EARNINGS
    def fill_actual_variance(self):
        first_col = 'B' # excel first column
        last_col = 'N' # excel last column 

        max_val = max(self.get_cell(name) for name in C('B31', 'N31')) # max value of Earnings Variance to Prior
        y0 = self.get_element_coords(self.get_element_by_title('actual-variance-min'))[1]
        max_h = (y0-self.get_element_coords(self.get_element_by_title('actual-variance-max'))[3])*2/3

        # set rectangles
        for name in C('B31', 'N31'):
            val = self.get_cell(name) # get excel value
            h = abs(max_h*val/max_val) # rectangle height
            rect = self.get_element_by_title(name + ' R')
            text = self.get_element_by_title(name)
            srgbClr = self.xpath('.//a:srgbClr', rect)[0]

            if val < 0:
                color = VARIANCE_NEG_COLOR # if Variance decrised, grey rectangle
                y = y0
                text_y = y0-self.get_element_sizes(text)[1]
            else:
                color = VARIANCE_POS_COLOR # if Variance encrised, blue rectangle
                y = y0-h
                text_y = y-self.get_element_sizes(text)[1]

            # set rectangle position, color, label
            self.set_element_size(rect, None, h)
            self.set_element_pos(rect, None, y)
            self.set_element_pos(text, None, text_y)

            srgbClr.set('val', color)

        # YTD Prior Var. Amount
        total_hp = self.get_cell('D33') # Height of total Variance
        total_label_e = self.E('B33') # Total YTD Variance
        total_rect_e = self.E('B33 R') # Total Variance rectangle
        total_original_h = self.get_element_sizes(total_rect_e)[1]
        total_h = max_h*total_hp
        self.set_element_size(total_rect_e, None, total_h) # set total variance rectangle size
        self.mod_element_pos(total_rect_e, None, total_original_h-total_h) # set total variance rectangle position
        # set label position
        total_label_e = self.set_element_pos(
            total_label_e,
            None,
            self.get_element_coords(total_rect_e)[1]-self.get_element_sizes(total_label_e)[1]
        )

        # underline current year
        current_year_line = self.E('actual-variance-current-year-line')
        x1 = self.get_element_coords(current_year_line)[2]
        period_values = [self.get_cell(c) for c in C('%s30' % first_col, '%s30' % last_col)] # enumerate current year values under rectangle

        # get previous year
        prev = None
        for year, col in zip(period_values, alpha_range(first_col, last_col)):
            if year == 1:
                break # current year
            prev = col # previous year

        # current year line begining
        first_cur_label_e = self.E('%s30' %col)
        if prev:
            x0 = (self.get_element_center(first_cur_label_e)[0]+self.get_element_center(self.E('%s30' % prev))[0])/2
        else:
            x0 = self.get_element_coords(first_cur_label_e)[0]

        # underline current year rectangles
        self.set_element_pos(current_year_line, x0, None)
        self.set_element_size(current_year_line, x1-x0, None)

    # Top 10 Managed Clients
    def fill_top10(self):
        separate_chart_group = self.E('separate-top-10-managed-clients')

        v = self.get_cell('B49') # max Axis scale
        values = [v*i/6 for i in range(6, 0, -1)]
        cells = ['B49'] + ['%s/6*B49' % i for i in range(5, 0, -1)] # Axis points between Max and Min
        # set text between Axis-max and Axis-min
        for v, c in zip(values, cells):
            self.set_text(c, str(int(round(v/10**6))))

        top10_max_e = self.get_element_by_title('top10-max')
        top10_max_coords = self.get_element_coords(top10_max_e)
        max_w = self.get_element_sizes(top10_max_e)[0] # Top scale according to the max value
        max_wp = 1.5 # YoY Rev. Growth Axis-max
        max_val = self.get_cell('B49') # Top scale
        min_wp = -.10 #YoY Rev. Growth Axis-min

        # setting Client YoY growth 
        for r in range(51, 61):
            revenue = self.get_cell('B%s' % r) # revenue empty rectangle
            margin = self.get_cell('C%s' % r) # margin solid rectangle
            yoy = self.get_cell('D%s' % r) # YOY growth arrow

            revenue_rect = self.get_element_by_title('revenue-%s' % r) # empty rectangle
            margin_rect = self.get_element_by_title('margin-%s' % r) # solid blue rectangle

            # set empty and solid rectangle size
            self.set_element_size(revenue_rect, revenue/max_val*max_w, None)
            self.set_element_size(margin_rect, margin/max_val*max_w, None)

            coords = self.get_element_coords(revenue_rect)
            x = coords[0]
            y = (coords[1]+coords[3])/2
            # If YoY % growth is negative, arrow goes left (to a max of -25%), and the amount is displayed above in parentheses 
            # If YoY growth is higher than 150%, display “line zig zag”, arrowhead at a value of 130%, and the amount above the arrow.
            # set arrow length
            if yoy <= max_wp:
                # ordinar arrow
                if yoy > 0:
                    # positive arrow
                    l = yoy*max_w/max_wp
                else:
                    l = -max(yoy, min_wp)*max_w/max_wp
                arrow = self.clone_template('top-10-arrow')
                self.set_element_size(arrow, l, None) # set arrow length
                if yoy < 0:
                    # negative arrow
                    self.xpath('.//a:xfrm', arrow)[0].attrib['flipH'] = '1' # arrow horizontal flip
                    x -= l
                    label = self.clone_template('top-10-arrow-label')
                    self.set_element_pos(
                        label,
                        top10_max_coords[0]-self.get_element_sizes(label)[0],
                        y-self.get_element_sizes(label)[1],
                    )
                    separate_chart_group.append(label)
                    self.set_element_text(label, '('+self.format_percent(yoy, prec=0)+'%)')
            else:
                # zig zag arrow
                arrow = self.clone_template('top-10-arrow-max')
                label = self.clone_template('top-10-arrow-label')
                self.set_element_pos(
                    label,
                    top10_max_coords[2],
                    y-self.get_element_sizes(label)[1],
                )
                separate_chart_group.append(label)
                self.set_element_text(label, self.format_percent(yoy, prec=0)+'%')
                y -= self.get_element_sizes(arrow)[1]/2

            self.set_element_pos(arrow, x, y)
            separate_chart_group.append(arrow)

    # Headline Metrics 
    def fill_headline(self):
        def flip(e):
            # set arrow flip: up or down
            xfrm = self.xpath('.//a:xfrm', e)[0]
            xfrm.set('flipV', '1')

        h5_val = self.get_cell('H5')
        if h5_val > 0:
            h5_label = self.get_element_by_title('H5') # set headline title
            h5_cat = self.get_element_by_title('H5-cat') # set headline category
            h5_arrow = self.get_element_by_title('H5 Arrow') # set headline arrow
            flip(h5_arrow)

            # set elements on template
            self.mod_element_pos(h5_arrow, None, -self.get_element_sizes(h5_cat)[1])
            self.mod_element_pos(h5_label, None, self.get_element_sizes(h5_arrow)[1])
            self.mod_element_pos(h5_cat, None, self.get_element_sizes(h5_arrow)[1])

    def fill_values(self):
        super(Dashboard7, self).fill_values()

        self.fill_headline()
        self.fill_blue_rect()
        self.fill_rates_by_geography()
        self.fill_key_metrics()
        self.fill_actual_variance()
        self.fill_top10()


if __name__ == '__main__':
    CMDHandler(
        Dashboard7,
    )
