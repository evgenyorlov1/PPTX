#!/usr/bin/env python

from __future__ import unicode_literals

import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, alpha_range, CMDHandler


VARIANCE_NEG_COLOR = '938B95'
VARIANCE_POS_COLOR = '0A9FDA'


def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in xrange(0, len(l), n):
        yield l[i:i+n]


class Dashboard10(PPTXGenerator):

    sheet_name = 'Margin Dashboard'
    template_shape_names = (
        'utilization-label',
        'key-metrics-label',
    )
    separate_charts = {
        'separate-headline-metrics': '10-Headline-Metrics.pptx',
        'separate-rates-cs-salary': '10-Rates-CS-Salary-Growth.pptx',
        'separate-utilization-headcount-growth': '10-Utilization-Headcount-Growth.pptx',
        'separate-earnings-per-partner': '10-Earnings-Per-Partner.pptx',
    }

    def get_simple_fillers(self):
        def big_money(v):
            mlns = round(float(v)/10**6, 0)
            if mlns > 999:
                return '$%sB' % self.format_float(mlns/10**3, prec=2, strip=False)
            else:
                return '$%sM' % self.format_float(mlns, prec=0)

        return {
            lambda v: self.format_float(v, prec=0): [
                C('B9', 'AN9'),
            ],
            lambda v: self.format_percent(v, prec=1)+'%': [
                C('B28', 'E28'),
            ],
            lambda v: self.format_percent(min(v, .999), prec=1)+'%': [
                C('B3'), C('D3'), C('F3'), C('H3'), C('J3'), C('L3'),
            ],
            lambda v: self.format_float(abs(v), prec=0)+'BPS': [
                C('B4'), C('D4'), C('F4'), C('H4'), C('J4'), C('L4'),
                C('H5'),
            ],
            lambda v: '$'+self.format_float(v/10.**3, prec=1)+'K': [
                C('B27', 'E27'),
            ],
            lambda v: '$'+self.format_float(v, prec=0): [
                C('B29', 'E29'),
                C('B31', 'E31'),
                C('B33', 'E33'),
            ],
            lambda v: self.format_float(v, prec=0): [
                C('B35', 'E35'),
            ],
            lambda v: self.format_float(v, prec=1): [
                C('B26', 'E26'),
            ],
        }

    # Rates vs. CS Salary Growth ok
    def fill_rates_and_salary(self):
        salary_chart = self.get_chart_by_title('salary-chart')
        salary_chart.fill_series(
            [self.get_cell(cell) for cell in C('B11', 'AN11')], 1,  # CS Salary per CS Staff
        ) 
        salary_chart_max = self.get_cell('H8') # Right Y Axis (Rate) Max 
        salary_chart.set_axis_max(self.format_float(salary_chart_max)) # set Right Y Axis Max
        salary_chart.set_axis_min(self.format_float(0)) # set Right Y Axis Min
        salary_chart.write()

        rate_chart = self.get_chart_by_title('rate-chart')
        rate_chart.fill_series(
            [self.get_cell(cell) for cell in C('B10', 'AN10')], 1, # Average Firm Rate
        )
        rate_chart_max = self.get_cell('F8') # Left Y Axis (Headcount) Max:
        rate_chart.set_axis_max(self.format_float(rate_chart_max)) # set Left Y Axis Max
        rate_chart.set_axis_min(self.format_float(0)) # set Left Y Axis Min
        rate_chart.write()

        # set Axis labels
        labels = [('F8*%s/5' % i, rate_chart_max*i/5.) for i in range(1, 5)]+[('F8', rate_chart_max)] # numerate Left Y Axis (Headcount)
        labels += [('H8*%s/5' % i, salary_chart_max*i/5.) for i in range(1, 5)]+[('H8', salary_chart_max)] # numerate Right Y Axis (Rate)
        labels += [('B12', self.get_cell('B12'))] # dotted line label
        for name, value in labels:
            self.set_text(name, '$'+self.with_comma(value)) 

        coords = self.get_plot_area_coords_from_chart('rate-chart')
        chart_h = coords[3]-coords[1]
        chart_h = coords[3]-coords[1]   

        # set Rolling rate Average line and label
        self.set_element_pos(
            self.get_element_by_title('rolling-39'),
            None,
            coords[3]-chart_h*(self.get_cell('B12')*1./rate_chart_max)
        )
        self.set_element_pos(
            self.get_element_by_title('B12'),
            None,
            coords[3]-chart_h*(self.get_cell('B12')*1./rate_chart_max)
        )

    # Utilization & Headcount Growth ok
    def fill_utilization_and_headcount(self):
        separate_chart_group = self.E('separate-utilization-headcount-growth')

        utilization_base = self.get_cell('B21') # Value at Origin (Y Axis)
        minyp = utilization_base-.1 
        maxyp = utilization_base+.1
        for title, value in zip(['B21-min', 'B21', 'B21-max'], [minyp, utilization_base, maxyp]):
            self.set_text(title, self.format_percent(value, prec=1)+'%')

        x0 = self.get_element_coords(self.get_element_by_title('headcount-min'))[0] # Headcount x-Axis x0
        x1 = self.get_element_coords(self.get_element_by_title('headcount-max'))[0] # Headcount x-Axis x1
        w = x1-x0 # chart width
        h = self.get_element_sizes(self.E('utilization-axis'))[1] # chart height
        y1 = self.get_element_coords(self.E('utilization-axis'))[3] # Utilization Axis(y-Axis)

        minp = -.05 # min x-Axis
        maxp = .2 # max x-Axis
        step = .005 # tick step

        max_size = self.get_element_sizes(self.E('utilization-B'))[0]
        for col in alpha_range('B', 'F'):
            utilization_e = self.E('utilization-%s' % col) # Utilization label(Current)
            utilization_prior_e = self.E('utilization-prior-%s' % col) # Utilization label(Prior)
            label_e = self.clone_template('utilization-label')
            for e, utilization_cell in zip([utilization_e, utilization_prior_e], ['%s18' % col, '%s19' % col]):
                size = max_size*self.get_cell('%s20' % col) # label size
                self.set_element_size(e, size, size)
                utilization = self.get_cell(utilization_cell)

                # tick position
                v = self.get_cell('%s17' % col)
                if v < minp:
                    # If HC Growth rate is below  -5%,  display at -5.5%
                    pos_v = minp-step
                elif v > maxp:
                    # If HC Growth rate is above 20%, display at 22.5%
                    pos_v = maxp+step
                else:
                    pos_v = v

                # set utilization label
                x = x0 + (pos_v-minp)/(maxp-minp)*w-size/2
                y = y1 - (utilization-minyp)/(maxyp-minyp)*h-size/2
                self.set_element_pos(e, x, y)

                # set Utilization label(Total HC, HC YoY, Utilization)
                if e is utilization_e:
                    # set label values
                    text_lines = [
                        'Total HC: %s' % self.with_comma(self.get_cell('%s16' % col)),
                        'HC YoY: %s%%' % self.format_percent(v, prec=1),
                        'Utilization: %s%%' % self.format_percent(utilization, prec=1),
                    ]
                    # set label text and position
                    self.set_element_text_lines(label_e, text_lines)
                    self.set_element_pos(
                        label_e,
                        x-self.get_element_sizes(label_e)[0],
                        y+(size-self.get_element_sizes(label_e)[1])/2,
                    )
                    separate_chart_group.append(label_e)

    # Earnings per partner
    def fill_earnings(self):
        max_y = float(self.get_cell('B38')) # Max of Y Axis (Earnings per CS Staff)
        max_x = float(self.get_cell('B39')) # Max of X Axis (Leverage W/ Partners)

        # numerate Y Axis
        for val, title in [(max_y*i/5., 'B38*%s/5' % i) for i in xrange(1, 5)]+[(max_y, 'B38')]:
            for e in self.get_elements_by_title(title):
                self.set_element_text(e, '$'+self.format_float(val/10.**3, prec=0)+'K')

        # numerate X Axis
        for val, title in [(max_x*i/5., 'B39*%s/5' % i) for i in xrange(1, 5)]+[(max_x, 'B39')]:
            for e in self.get_elements_by_title(title):
                self.set_element_text(e, self.format_float(val, prec=0))

        w = self.get_element_sizes(self.E('earnings-x-axis'))[0] # char width
        h = self.get_element_sizes(self.E('earnings-y-axis'))[1] # chart height

        for col in alpha_range('B', 'E'):
            rect_e = self.E('rect-%s' % col) # section rectangle(or section)

            original_h = self.get_element_sizes(rect_e)[1] # section rectangle orig. height 
            wpn = self.get_cell('%s26' % col)/max_x # section rect. width
            hpn = self.get_cell('%s25' % col)/max_y # section rect. height
            hn = hpn*h # section chart height
            # set section
            self.set_element_size(
                rect_e,
                w*wpn,
                hn,
            )
            self.mod_element_pos(
                rect_e,
                None,
                original_h-hn,
            )

            rect_coords = self.get_element_coords(rect_e)
            label_e = self.E('earnings-label-%s' % col) # section rectangle label

            label_y = None
            if hpn < .33:
                # If lower < 33%, Move EPP Value up so that the bottom is at 33% 
                label_y = rect_coords[3]-.33*h-self.get_element_sizes(label_e)[1]
            if wpn < .25:
                # if lower < 25%, Move text so that the the left edge is at 25% 
                label_x = rect_coords[0]+.25*w
            else:
                label_x = (rect_coords[0]+rect_coords[2]-self.get_element_sizes(label_e)[0])/2

            if hpn < .33 or wpn < .25:
                # if width < 25%, display text at the color of the function 
                # if height < 33%, display text at the color of the function 
                color = self.get_element_fill_color(rect_e)
                for e in self.xpath('.//a:rPr[a:solidFill]', label_e):
                    self.set_element_text_color(e, color)
                self.replace_pic_color(label_e, color)

            # set section label
            self.set_element_pos(
                label_e,
                label_x,
                label_y,
            )

    # Key Metrics ok
    def fill_key_metrics(self):
        separate_chart_group = self.E('separate-earnings-per-partner')

        x0 = self.get_element_coords(self.get_element_by_title('key-metrics-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('key-metrics-max'))[0]
        w = x1-x0 # chart width

        minp = -.10 # x Axis min (-10%)
        maxp = .15 # x Axis max (15%)
        step = .005 # tick step

        def set_tick(name):
            tick = self.get_element_by_title(name)

            v = self.get_cell(name)
            tick_h, tick_w = self.get_element_sizes(tick) # tick height, width
            show_label = False # label above tick(%)
            if v < minp:
                # If growth rates of metrics are lower than -10%
                pos_v = minp-step # display at 15.5% level
                show_label = True # display actual value above the dot in 7pt font 
            elif v > maxp:
                # If growth rates of metrics are higher than 15%
                pos_v = maxp+step # display at -10.5% level
                show_label = True # display actual value above the dot in 7pt font 
            else:
                pos_v = v

            x = x0 + (pos_v-minp)/(maxp-minp)*w # tick position
            self.set_element_pos(tick, x-tick_w/2, None) # set tick position

            # tick label value
            if -0.99 < v < 0.99:
                text = self.format_percent(v, prec=1)+'%'
            else:
                text = self.format_percent(min(max(v, -0.99), .99), prec=0)+'%' # Max value 99%; Min -99%

            # if label's value is )-10;15(
            if show_label:
                value_e = self.clone_template('key-metrics-label')
                value_e_w, value_e_h = self.get_element_sizes(value_e) # tick label's width, heigh
                label_x = x-value_e_w/2 # tick label's x 
                label_y = self.get_element_coords(tick)[1]-value_e_h # # tick label's y
                self.set_element_text(value_e, text) # set tick label's text
                self.set_element_pos(value_e, label_x, label_y) # set tick label's position
                separate_chart_group.append(value_e)

        for n in [30, 32, 34, 36, 37]:
            for col in alpha_range('B', 'E'):
                set_tick('%s%s' % (col, n))

    def fill_values(self):
        super(Dashboard10, self).fill_values()

        self.fill_rates_and_salary()
        self.fill_utilization_and_headcount()
        self.fill_earnings()
        self.fill_key_metrics()


if __name__ == '__main__':
    CMDHandler(
        Dashboard10,
    )
