#!/usr/bin/env python

from __future__ import unicode_literals

from collections import defaultdict
import sys
from os import path

from lxml import etree

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, CMDHandler


VARIANCE_NEG_COLOR = '938B95' # hex grec
VARIANCE_POS_COLOR = '0A9FDA' # hex blue 


def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in xrange(0, len(l), n):
        yield l[i:i+n]


class Dashboard9(PPTXGenerator):

    sheet_name = 'Firm Dashboard'
    template_shape_names = (
        'marg-arrow-up',
        'marg-arrow-down',
        'tick-label',
        'icon-advisory',
        'icon-audit',
        'icon-consulting',
        'icon-tax',
        'rev-earn-label',
    )
    separate_charts = {
        'separate-headline-metrics': '9-Headline-Metrics.pptx',
        'separate-revenue-and-earnings': '9-Revenue-Earnings-by-Major-Service-Area.pptx',
        'separate-cost-breakdown': '9-COST-BREAKDOWN.pptx',
        'separate-revenue-and-earnings-variance': '9-Revenue-Earnings-Variance-to-Plan.pptx',
        'separate-key-metrics': '9-Key-Metrics.pptx',
        'separate-revenue': '9-Revenue-Variance-to-Plan.pptx',
        'separate-earnings': '9-Earnings-Variance-to-Plan.pptx',
    }

    def get_simple_fillers(self):
        def big_money(v):
            mlns = round(float(v)/10**6, 0)
            if mlns > 999:
                return '$%sB' % self.format_float(mlns/10**3, prec=2, strip=False)
            else:
                return '$%sM' % self.format_float(mlns, prec=0)

        return {
            lambda v: v: [
                C('B21', 'F21'),
                C('A37', 'A41'),
            ],
            lambda v: self.format_float(v, prec=0): [
                C('B46', 'N46'),
                C('B53', 'N53'),
            ],
            lambda v: self.format_float(v/10.**6, include_sign=True): [
                C('B47', 'N47'),
                C('B54', 'N54'),
            ],
            big_money: [
                C('B3'), C('D3'),
                C('B4'), C('D4'),
            ],
            lambda v: self.format_percent(v, prec=1, strip=False)+'%': [
                C('F3'),
                C('B5'), C('D5'), C('H5'),
            ],
            lambda v: self.format_float(v, prec=0)+'BPS': [
                C('F4'),
                C('F5'),
            ],
            lambda v: self.format_float(abs(v), prec=0): [
                C('H4'),
            ],
            self.with_comma: [
                C('H3'),
            ],
            lambda v: '$'+self.with_comma(v/10.**6): [
                C('B8'),
            ]
        }

    # KEY METRICS ok
    def fill_key_metrics(self):
        separate_chart_group = self.E('separate-key-metrics') # get Key Metrics chart

        x0 = self.get_element_coords(self.get_element_by_title('key-metrics-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('key-metrics-max'))[0]
        w = x1-x0 # chart width

        # Scale is constant at -10% to 15%
        minp = -.10 # x-Axis min
        maxp = .15 # x-Axis max
        step = .005 # tick step

        def set_tick(name, show_label, level, min_stacked, max_stacked):
            tick = self.get_element_by_title('tick-'+name)

            v = self.get_cell(name)
            tick_h, tick_w = self.get_element_sizes(tick) # tick height & width
            if v < minp:
                # If  amount is lower than -10%:
                # Display at -10.5%
                pos_v = minp-step
            elif v > maxp:
                # If amount is higher than 15%
                # Display at 15.5%
                pos_v = maxp+step
            else:
                pos_v = v

            x = x0 + (pos_v-minp)/(maxp-minp)*w # tick position
            self.set_element_pos(tick, x-tick_w/2, None)

            shift = -tick_h/3*level if min_stacked or max_stacked else 0 # shift up label if stacked
            self.mod_element_pos(tick, None, shift)

            # if Lower than -99%, just display -99%
            #  if higher than 99%, just display 99%
            if -0.99 < v < 0.99:
                text = self.format_percent(v, prec=1)+'%'
            else:
                text = self.format_percent(min(max(v, -0.99), .99), prec=0)+'%'

            # label above tick
            if show_label:
                value_e = self.clone_template('tick-label')
                value_e_w, value_e_h = self.get_element_sizes(value_e) # tick label height and width
                # label x coord
                if min_stacked:
                    label_x = x
                elif max_stacked:
                    label_x = x-value_e_w
                else:
                    label_x = x-value_e_w/2

                # label y coord
                if min_stacked or max_stacked:
                    label_y = self.get_element_coords(tick)[1]-(tick_h-value_e_h)/2+shift 
                else:
                    label_y = self.get_element_coords(tick)[1]-value_e_h
                # set tick label position and text
                self.set_element_text(value_e, text)
                self.set_element_pos(value_e, label_x, label_y)
                separate_chart_group.append(value_e)

        # set all ticks
        for n in range(37, 41):
            cols = 'BCDEF' # excel cols
            cells = ['%s%s' % (c, n) for c in cols]
            values = [self.get_cell(c) for c in cells] # tick value
            # If multiple icons are either above 15% or below -10% on the same metric, stack them vertically
            min_stacked = sum(v < minp for v in values[:-1]) > 1
            max_stacked = sum(v > maxp for v in values[:-1]) > 1
            min_level = -1
            max_level = -1
            for col, v in zip(cols, values):
                level = 0 # number of stacked
                if v < minp:
                    level = min_level = min_level+1
                if v > maxp:
                    level = max_level = max_level+1
                cell = '%s%s' % (col, n)
                # set tick on chart
                set_tick(
                    cell,
                    col != 'F' and (v < minp or v > maxp),
                    level,
                    min_stacked and v < minp,
                    max_stacked and v > maxp,
                )

    # Firm Earnings Variance to Prior ok
    def fill_firm_earnings(self):
        cells = C('B54', 'N54') # excel values
        max_val = max(abs(self.get_cell(name)) for name in cells) # max value for scale
        y0 = self.get_element_coords(self.get_element_by_title('firm-earnings-min'))[1]
        max_h = abs(y0-self.get_element_coords(self.get_element_by_title('firm-earnings-max'))[3])
        y0 -= max_h
        max_h = max_h*2./3 # Max amount of the 13 variances is 2/3 height between x axis and bottom of “Actual Variance to Prior” text

        for name in cells:
            val = self.get_cell(name) # cell value
            h = abs(max_h*val/max_val) # rectangle height
            rect = self.get_element_by_title(name + ' R') # rectangle
            text = self.get_element_by_title(name) # rectangle label
            srgbClr = self.xpath('.//a:srgbClr', rect)[0] # rectangle color

            if val < 0:
                color = VARIANCE_NEG_COLOR # if amount is < 0, grey color
                y = y0
                # if amounts are negative, display just above x-axis
                text_y = y0-self.get_element_sizes(text)[1]
            else:
                color = VARIANCE_POS_COLOR # if amount is > 0, blue color
                y = y0-h
                # Display amounts just above bars
                text_y = y-self.get_element_sizes(text)[1]

            # set rectangle and text
            self.set_element_size(rect, None, h)
            self.set_element_pos(rect, None, y)
            self.set_element_pos(text, None, text_y)

            # set rectangle color
            srgbClr.set('val', color)

        # set Prior Var. Amount
        year = self.get_element_by_title('D56') # Height of Total Variance (% of Max) 
        year_value = self.get_element_by_title('B56') # Total YTD Variance
        h = self.get_cell('D56')*max_h # height of Total Varience 
        y = self.get_element_coords(self.get_element_by_title('firm-earnings-min'))[1]-h 
        self.set_element_size(year, None, h) # set Total Varience size
        self.set_element_pos(year, None, y) # set Total Varience position

        # set Prior Var. Amount label text and position
        self.set_element_text(year_value, '$'+self.format_float(self.get_cell('B56')/10.**6, prec=0))
        self.set_element_pos(year_value, None, y-self.get_element_sizes(year_value)[1])

    # Revenue Variance to Prior ok
    def fill_firm_revenue(self):
        cells = C('B47', 'N47') # excel values
        max_val = max(abs(self.get_cell(name)) for name in cells) # rectangle height
        y0 = self.get_element_coords(self.get_element_by_title('firm-revenue-min'))[1]
        max_h = abs(y0-self.get_element_coords(self.get_element_by_title('firm-revenue-max'))[3])
        y0 -= max_h
        max_h = max_h*2./3 # Max amount of the 13 variances is 2/3 height between x axis and bottom of “Actual Variance to Prior” text

        for name in cells:
            val = self.get_cell(name) # cell value
            h = abs(max_h*val/max_val) # rectangle height
            rect = self.get_element_by_title(name + ' R') # rectangle
            text = self.get_element_by_title(name) # rectangle label

            if val < 0:
                y = y0
                # if amounts are negative, display just above x-axis
                text_y = y0-self.get_element_sizes(text)[1]
            else:
                y = y0-h
                # Display amounts just above bars
                text_y = y-self.get_element_sizes(text)[1]

            # set rectangle and text
            self.set_element_size(rect, None, h)
            self.set_element_pos(rect, None, y)
            self.set_element_pos(text, None, text_y)

        # # set Prior Var. Amount
        year = self.get_element_by_title('D49') # Height of Total Variance (% of Max)
        year_value = self.get_element_by_title('B49') # Total YTD Variance
        h = self.get_cell('D49')*max_h # height of Total Varience 
        y = self.get_element_coords(self.get_element_by_title('firm-revenue-min'))[1]-h
        self.set_element_size(year, None, h) # set Total Varience size
        self.set_element_pos(year, None, y) # set Total Varience position

        # set Prior Var. Amount label text and position
        self.set_element_text(year_value, '$'+self.format_float(self.get_cell('B49')/10.**6, prec=0))
        self.set_element_pos(year_value, None, y-self.get_element_sizes(year_value)[1])

    # REVENUE & EARNINGS BY MAJOR SERVICE AREA 
    def fill_blue_rect(self):
        separate_chart_group = self.E('separate-revenue-and-earnings')

        coords = self.get_element_coords(self.shapes['big-blue-rect'])
        w, h = coords[2]-coords[0], coords[3]-coords[1] # chart widht and height
        x0, y0 = coords[0], coords[3]
        # rectangle sizes and values
        sizes = zip(
            'BCDE',
            [self.get_cell(c) for c in C('B12', 'E12')], # rectangle width(%)
            [self.get_cell(c) for c in C('B18', 'E18')], # CE Height(%) 
            [self.get_cell(c) for c in C('B14', 'E14')], # EBA Height (Dotted Line)(%)
        )
        # f width of section is below 4% Do not display. Add width to next smallest section to maintain size.
        sizes = [(l, wn, hn, lpn) if wn > 0.04 else (l, 0., 0.) for l, wn, hn, lpn in sizes]
        norm = sum(wn for l, wn, hn, lpn in sizes)

        x = x0
        for l, wpn, hpn, lpn in sizes:
            wn = wpn*w/norm # normilize width

            hpn_min, hpn_max = .2, .69 # Min height of CE Height is 20%; Max height of CE Height is 69%
            lpn_max = .77 # The Maximum height of the dotted line (EBA Height) is 77%

            if wpn > .10:
                # If Width of a section is above 10% of total width
                def format_money(v):
                    # Max value of all $ figures is $9999
                    if v > 9999:
                        # If higher put as $1X.XB (only 1 decimal point)    
                        return '$%sB' % self.format_float(v/10**3, prec=1)
                    return '$%s' % self.format_float(v, prec=0)
                font_size = 15 # Font size of all values is 15 pt
            elif wpn > .07:
                # If Width of a Section is between 10%-7% of total width
                def format_money(v):
                    # If a $ value reaches 4 digits, display in Billions with 1 decimal place($7894 = $7.9B)
                    if v > 999:
                        return '$%sB' % self.format_float(v/10**3, prec=1)
                    return '$%s' % self.format_float(v, prec=0)
                font_size = 12 # Change font size of all values to 12pt
            elif wpn > .04:
                # If width of a section is between 7%-4%
                hpn_min, hpn_max = 0, .9 # No MIN CE Height value; Max CE Height is 90%
                lpn_max = .91 # MAX of Dotted line is 91%
            else:
                # If width of section is below 4%
                # Do not display.
                hpn = 0

            hn = min(max(hpn, hpn_min), hpn_max)*h if wpn else 0 # CE Height
            ln = min(lpn, lpn_max)*h if wpn else 0 # EBA Height

            rect = self.shapes['%s-rect' % l] # section rectangle
            line = self.get_element_by_title('%s-line' % l) # dotted line

            revenue = self.get_element_by_title('%s10' % l) # REV value
            revenue_arrow = self.get_element_by_title('%s11-arrow' % l) # Revenue Growth From Prior
            earnings = self.get_element_by_title('%s13' % l) # EBA
            ce_marg = self.get_element_by_title('%s15' % l) # Controllable Earnings
            ce_marg_p = self.get_element_by_title('%s16' % l) # CE Margin

            if wpn <= .07:
                # If width of a section is between 7%-4%
                # Remove all values, just display correct heights of margin boxes/lines
                for e in [revenue, revenue_arrow, earnings, ce_marg, ce_marg_p]:
                    remove(e)
            else:
                # set section value labels
                for e in [revenue, earnings, ce_marg, ce_marg_p]:
                    self.set_element_text_size(e, font_size)

                # section icon(Ad, A, C and T)
                icon = self.clone_template('icon-%s' % self.get_cell('%s9' % l).lower())
                self.set_element_pos(icon, x, None)
                separate_chart_group.append(icon)

                # set Rev value
                self.set_element_pos(revenue, x, None)
                self.set_element_text(revenue, format_money(self.get_cell('%s10' % l)/10.**6))

                # set revenue arrow
                revenue_arrow_shift = self.get_element_sizes(revenue)[0]-self.get_element_sizes(revenue_arrow)[0]
                self.set_element_pos(revenue_arrow, x+revenue_arrow_shift/2, None)
                self.set_element_text(revenue_arrow, self.format_percent(self.get_cell('%s11' % l), prec=0)+'%')

                # set EBA
                earnings_value = self.get_cell('%s13' % l)/10.**6
                self.set_element_pos(earnings, x, y0-ln)
                self.set_element_text(earnings, format_money(earnings_value))

                # Controllable Earnings (C.E)
                ce_marg_val = self.get_cell('%s15' % l)/10.**6
                self.set_element_pos(ce_marg, x, None)
                self.set_element_text(ce_marg, format_money(ce_marg_val))

                # CE Margin (C.E Marg)
                ce_marg_p_val = self.get_cell('%s16' % l)
                self.set_element_pos(ce_marg_p, x, None)
                self.set_element_text(ce_marg_p, self.format_percent(ce_marg_p_val, prec=0)+'%')

                # set CE Margin arrow
                marg_arrow_val = self.get_cell('%s17' % l)
                marg_arrow = self.clone_template('marg-arrow-up' if marg_arrow_val >= 0 else 'marg-arrow-down')
                shift = self.get_element_sizes(ce_marg_p)[0]-self.get_element_sizes(marg_arrow)[0]
                self.set_element_text(marg_arrow, self.format_float(min(abs(marg_arrow_val), 999), prec=0)+'BPS') 
                self.set_element_pos(marg_arrow, x+shift/2, None)
                separate_chart_group.append(marg_arrow)

            if wpn == 0:
                # If width is 0, remove section
                remove(rect)
                remove(line)
            else:
                # set rectangle
                self.set_element_pos(rect, x, y0-hn)
                self.set_element_size(rect, wn, hn)

                # set dotted line
                self.set_element_pos(line, x, y0-ln)
                self.set_element_size(line, wn, None)

            x += wn

            if l != 'E':
                s = self.get_elements_by_title('%s-sep' % l)[0]
                if wn == 0:
                    remove(s)
                else:
                    self.set_element_pos(s, x, None)

    # COST BREAKDOWN ok
    def fill_cost_breakdown(self):
        def rotate(e):
            # rotate on 90 deg
            self.xpath('.//a:xfrm', e)[0].set('rot', '-5400000')
            w, h = self.get_element_sizes(e)
            shift = (h-w)/2
            coords = self.get_element_coords(e)
            x, y = coords[0], coords[1]
            x += shift
            y -= shift
            self.set_element_pos(e, x, y)
            try:
                pPr = self.xpath('.//a:pPr', e)[0]
            except ValueError:
                pPr = etree.Element('{%s}pPr' % self.NS['a'])
                self.xpath('.//a:p', e)[0].insert(0, pPr)
            pPr.set('algn', 'r')

        cols = 'BCDEF'

        values = [] # (rec. width, amount($))
        s = 0
        for l in cols:
            amount = round(self.get_cell('%s22' % l)/10.**6, 0) # Cerrent Amount
            width = self.get_cell('%s24' % l) # % of Max

            # If lower than 3% of total width do not display
            # If Value is $10,000+ and width below 7.5% do not display
            if width < .03 or amount >= 10**4 and width < .075:
                width = 0

            values.append((width, amount))
            s += width # real width

        values = [(w/s, a) for w, a in values] # change values for new width
        marg = 12700*2

        box = self.get_element_by_title('cost-breakdown-box') # chart
        coords = self.get_element_coords(box)
        x0, y0 = coords[0], coords[1]
        w, h = self.get_element_sizes(box) # chart width and height
        x = x0
        for col, (wn, an) in zip(cols, values):
            title = self.get_element_by_title('%s21' % col) # rectangle title(CS Sal. P&A etc)
            amount = self.get_element_by_title('%s22' % col) # amount
            yoy = self.get_element_by_title('%s23-arrow' % col) # YOY growth (arrow) 
            rect = self.get_element_by_title('%s24-rect' % col) # rectangle width
            color = self.get_cell('%s25' % col) # rectangle color
            color = '93C83D' if color.strip().lower() == 'green' else 'B1B3B5' # rectangle hex color

            # if width == 0, remove section 
            if wn == 0:
                for e in [title, amount, yoy, rect]:
                    remove(e)
                continue

            xn = wn*w

            def format_money(val):
                val = self.format_float(val, prec=0)
                if len(val) > 3:
                    val = val[:-3] + ',' + val[-3:]
                return '$' + val
            font_size = 14 # Font size for all figures is 14pt
            rotated = False
            # If value is less than $999
            if an < 999:
                # If between 7% and 3% width
                if wn < .07:
                    font_size = 12 # Change to 12pt font
                    rotated = True # Turn 90 degrees
            # If value is between $1,000-$9,999
            elif an < 9999:
                # If width is between 9% and 7.5%
                if .075 <= wn < .09:
                    font_size = 12 # display horizontally at 12 pt font
                # If between 7.5% and 3%
                elif wn < .075:
                    font_size = 10 # Change to 10pt font
                    rotated = False # Turn 90 degrees
            else:
                # If Value is $10,000+ (Max is $99,999)
                if .075 <= wn < .1:
                    # If between 10% and 7.5%
                    format_money = lambda val: '$' + self.format_float(val/1000, prec=1) + 'B' # isplay in billions with 1 decimal point
                    font_size = 12 # display at 12 pt font

            # set title position
            self.set_element_pos(title, x, y0-2*h)
            self.set_element_size(title, xn, 2*h)
            self.set_element_text_direction(title, 'vert270' if rotated else None) # set rotation
            if rotated:
                self.set_element_text(title, self.get_cell('%s21' % col)[:7])
                bodyPr = self.xpath('.//a:bodyPr', title)[0]
                bodyPr.set('lIns', '0')
                bodyPr.set('rIns', '0')
            else:
                lines = [''.join(c) for c in chunks(self.get_cell('%s21' % col), 7)][:3] # underline text
                self.set_element_text_lines(title, lines)

            # set amount
            self.set_element_pos(amount, x, y0)
            self.set_element_size(amount, xn, h)
            self.set_element_text(amount, format_money(an))
            self.set_element_text_size(amount, font_size)
            self.set_element_text_direction(amount, 'vert270' if rotated else None)

            # set YOY growth
            self.set_element_pos(yoy, x+(xn-self.get_element_sizes(yoy)[0])/2, None)
            self.set_element_text(yoy, self.format_percent(self.get_cell('%s23' % col), prec=1)+'%')

            # set rectangle
            self.set_element_pos(rect, x+marg, None)
            self.set_element_size(rect, xn-2*marg, None)
            srgbClr = self.xpath('.//a:srgbClr', rect)[0]
            srgbClr.set('val', color)

            x += xn

    # Revenue & Earnings – Actual Variance to Plan ok
    def fill_revenue_and_earnings(self):
        separate_chart_group = self.E('separate-revenue-and-earnings-variance')

        x0 = self.get_element_coords(self.get_element_by_title('rev-earn-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('rev-earn-max'))[0]
        w = x1-x0 # chart width

        # Scale is Constant (-5% to 15%)
        minp = -.05 # chart x-Axis min
        maxp = .15 # chart x-Axis max
        step = .005 # tick step

        positions = defaultdict(dict)

        def set_tick(i, below):
            name = '%s%s' % i
            tick = self.get_element_by_title('icon-'+name)
            color = {
                'B': '898B8D', # grey hex color
                'C': '1E2556', # blue hex color
                'D': '39A0DA', # light blue hex color
                'E': '43913E', # green hex color
            }.get(i[0])

            v = self.get_cell(name)
            tick_w = self.get_element_sizes(tick)[0]
            if v < minp: # If Value is below -5% display at -5.5%
                pos_v = minp-step
            elif v > maxp: # If value is above 15% display at 15.5%
                pos_v = maxp+step
            else:
                pos_v = v

            x = x0 + (pos_v-minp)/(maxp-minp)*w # tick position
            xend = x+tick_w
            self.set_element_pos(tick, x-tick_w/2, None) # set ticket on chart
            level = 0
            for (start, end), l in positions[below].iteritems():
                if start <= xend and x <= end:
                    level = max(level, l+1)
            positions[below][(x, xend)] = level

            text = '$'+self.format_float(self.get_cell('%s%s' % (i[0], i[1]+1))/10.**6, prec=0) # format values

            # earn labels
            value_e = self.clone_template('rev-earn-label')
            value_e_w, value_e_h = self.get_element_sizes(value_e) # earn label's width and height

            self.set_element_text(value_e, text)
            if color:
                self.set_element_text_color(value_e, color)
            level_shift = 0 if i[0] == 'F' else (1 if below else -1)*level*value_e_h # x-Axis level
            # set tick position
            self.set_element_pos(
                value_e,
                x-value_e_w/2,
                (self.get_element_coords(tick)[3]
                    if below else (self.get_element_coords(tick)[1]-value_e_h))+level_shift
            )
            separate_chart_group.append(value_e)

        # set ticks on chart
        for n in [29, 31]:
            cols = 'BCDEF'
            cells = ['%s%s' % (c, n) for c in cols] # excel cells
            values = [self.get_cell(c) for c in cells]
            for col, v in zip(cols, values):
                set_tick((col, n), n == 31)

    def fill_values(self):
        super(Dashboard9, self).fill_values()

        self.fill_key_metrics()
        self.fill_firm_earnings()
        self.fill_firm_revenue()
        self.fill_blue_rect()
        self.fill_cost_breakdown()
        self.fill_revenue_and_earnings()


if __name__ == '__main__':
    CMDHandler(
        Dashboard9,
    )
