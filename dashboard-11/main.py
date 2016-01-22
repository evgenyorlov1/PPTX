#!/usr/bin/env python

from __future__ import unicode_literals

from collections import defaultdict
import sys
from os import path

if __name__ == '__main__' and __package__ is None:
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

from data2ppt import C, PPTXGenerator, remove, alpha_range, CMDHandler


class Dashboard11(PPTXGenerator):

    sheet_name = 'Current Period Dashboard'
    template_shape_names = (
        'marg-arrow-up',
        'marg-arrow-down',
        'tick-label',
        'icon-advisory',
        'icon-audit',
        'icon-consulting',
        'icon-tax',
        'rev-earn-label',
        'revenue-arrow-up',
        'revenue-arrow-down',
    )
    separate_charts = {
        'separate-headline-metrics': '11-Headline-Metrics.pptx',
        'separate-revenue-and-earnings': '11-Revenue-and-Earnings-by-Business-and-Firm.pptx',
        'separate-revenue-and-earnings-variance': '11-Revenue-Earnings-Variance-to-Plan.pptx',
        'separate-key-metrics': '11-Key-Metrics.pptx',
        'separate-enabling-areas': '11-Enabling-Areas-Parent-Costs.pptx',
    }

    def get_simple_fillers(self):
        def big_money(v):
            mlns = round(float(v)/10**6, 0)
            if mlns > 999:
                return '$%sB' % self.format_float(mlns/10**3, prec=2, strip=False)
            else: 
                return '$%sM' % self.format_float(mlns, prec=0)

        return {
            lambda v: big_money(min(v, 99.99*10**9)): [
                C('B3'), C('D3'),
                C('B4'), C('D4'),
            ],
            lambda v: self.format_percent(min(v, .999), prec=1, strip=False)+'%': [
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
            lambda v: self.with_comma(min(v, 999999)): [
                C('H3'),
            ],
            lambda v: '$'+self.with_comma(v/10.**6): [
                C('B8'),
                C('B40', 'K40'),
            ],
            lambda v: 'Total $%s' % self.with_comma(v/10.**6): [
                C('B38')
            ],
            lambda v: 'Plan Var: $%s' % self.with_comma(v/10.**6): [
                C('F38'),
            ],
            lambda v: 'YoY Growth: %s%%' % self.format_percent(v, prec=1, strip=False): [
                C('D38'),
            ],
        }

    # Key Metrics ok
    def fill_key_metrics(self):
        separate_chart_group = self.E('separate-key-metrics')

        x0 = self.get_element_coords(self.get_element_by_title('key-metrics-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('key-metrics-max'))[0]
        w = x1-x0 # chart width

        # Scale is constant at -10% to 15% 
        minp = -.10 # scale Min
        maxp = .15 # scale Max
        step = .005 # tick step 

        def set_tick(name, show_label, level, min_stacked, max_stacked, color):
            # set tick on chart
            tick = self.get_element_by_title('tick-'+name)

            v = self.get_cell(name) # tick value
            tick_h, tick_w = self.get_element_sizes(tick) # tick height, width
            if v < minp:
                # If  amount is lower than -10%
                pos_v = minp-step # Display at -10.5% on scale
            elif v > maxp:
                # If amount is higher than 15% 
                pos_v = maxp+step # Display at 15.5% on scale
            else:
                pos_v = v

            x = x0 + (pos_v-minp)/(maxp-minp)*w # tick x position
            self.set_element_pos(tick, x-tick_w/2, None)

            shift = -tick_h/3*level if min_stacked or max_stacked else 0
            self.mod_element_pos(tick, None, shift)

            # if Lower than -99%, just display -99% 
            # if higher than 99%, just display 99%
            if -0.99 < v < 0.99:
                text = self.format_percent(v, prec=1)+'%'
            else:
                text = self.format_percent(min(max(v, -0.99), .99), prec=0)+'%'

            # set label on chart
            if show_label:
                value_e = self.clone_template('tick-label')
                value_e_w, value_e_h = self.get_element_sizes(value_e) # tick width, height
                if min_stacked:
                    # Move figures to the right of the Chiclet if below -5% and stacked
                    label_x = x+tick_w/2
                elif max_stacked:
                    # Move figures to the left of the Chiclet if above 15% and stacked 
                    label_x = x-value_e_w-tick_w/2
                else:
                    label_x = x-value_e_w/2

                if min_stacked or max_stacked:
                    label_y = self.get_element_coords(tick)[1]-(tick_h-value_e_h)/2+shift # label y Axis
                else:
                    label_y = self.get_element_coords(tick)[1]-value_e_h
                self.set_element_text(value_e, text)
                # set label color
                if color:
                    self.set_element_text_color(value_e, color)
                self.set_element_pos(value_e, label_x, label_y)
                separate_chart_group.append(value_e)

        # set ticks
        for n in range(31, 36):
            cols = 'BCDEF' # excel cols
            cells = ['%s%s' % (c, n) for c in cols] # excel cells
            values = [self.get_cell(c) for c in cells] # excel cell's values
            # If multiple icons are either above 15% or below -10% on the same metric, stack them 
            min_stacked = sum(v < minp for v in values[:-1]) > 1
            max_stacked = sum(v > maxp for v in values[:-1]) > 1
            min_level = -1
            max_level = -1
            # set tick color
            for col, v in zip(cols, values):
                color = {
                    'B': '898B8D', # Advisory
                    'C': '1E2556', # Audit
                    'D': '39A0DA', # Consulting
                    'E': '43913E', # Tax
                }.get(col)
                level = 0 # chart level
                if v < minp:
                    level = min_level = min_level+1
                if v > maxp:
                    level = max_level = max_level+1
                cell = '%s%s' % (col, n)
                # set tick
                set_tick(
                    cell,
                    col != 'F' and (v < minp or v > maxp),
                    level,
                    min_stacked and v < minp,
                    max_stacked and v > maxp,
                    color,
                )

    # Vertical Tree Map (top left) ok
    def fill_blue_rect(self):
        separate_chart_group = self.E('separate-revenue-and-earnings')

        coords = self.get_element_coords(self.shapes['big-blue-rect']) # base rectangle
        w, h = coords[2]-coords[0], coords[3]-coords[1] # rectangle width, height
        x0, y0 = coords[0], coords[3]
        # section sizes and values
        sizes = zip(
            'BCDE',
            [self.get_cell(c) for c in C('B12', 'E12')], # Width (%)
            [self.get_cell(c) for c in C('B18', 'E18')], # CE Height (%, if width != 0)
            [self.get_cell(c) for c in C('B14', 'E14')], # EBA Height (Dotted Line), if width != 0
        )
        # If width of section is below 4% Do not display
        sizes = [(l, wn, hn, lpn) if wn > 0.04 else (l, 0., 0.) for l, wn, hn, lpn in sizes]
        norm = sum(wn for l, wn, hn, lpn in sizes) # normalize section width

        x = x0
        for l, wpn, hpn, lpn in sizes:
            wn = wpn*w/norm # normalized section width

            gap_size = 0.08 # minimum difference between CE Margin height and EBA Margin Height is 8%
            hpn_min, hpn_max = .2, .69 # Min CE Height; Max CE Height
            lpn_min, lpn_max = .28, .77 # Min EBA Height; Max EBA Height

            if wpn > .10:
                # If Width of a section is above 10%
                def format_money(v): # Max value of all $ figures is $9999 
                    if v > 9999:
                        return '$%sB' % self.format_float(v/10**3, prec=1)
                    return '$%s' % self.format_float(v, prec=0)
                font_size = 15 # Font size of all values is 15 pt 
            elif wpn > .07:
                # If Width of a Section is between 10%-7%
                def format_money(v): # Max value of all $ figures is $9999 
                    if v > 999:
                        return '$%sB' % self.format_float(v/10**3, prec=1)
                    return '$%s' % self.format_float(v, prec=0)
                font_size = 12 # Change font size of all values to 12pt 
            elif wpn > .04:
                # If width of a section is between 7%-4% 
                hpn_min, hpn_max = 0, .9 # No MIN value; MAX of Blue square 90% 
                lpn_max = .91 # MAX of Dotted line is 91%
            else:
                # if width == 0, remove section
                hpn = 0

            gape_needed_adjustment = max(gap_size - (lpn - hpn), 0) # diff between CE Margin h. and EBA Margin h.
            if gape_needed_adjustment:
                # To fit text, minimum difference between CE Margin height and EBA Margin Height is 8% 
                # If difference is less than 8%:
                if lpn_max - lpn < gape_needed_adjustment/2:
                    hpn -= gape_needed_adjustment - (lpn_max - lpn)
                    lpn = lpn_max
                elif hpn - hpn_min < gape_needed_adjustment/2:
                    lpn += gape_needed_adjustment - (hpn - hpn_min)
                    hpn = hpn_min
                else:
                    hpn -= gape_needed_adjustment/2
                    lpn += gape_needed_adjustment/2

            hn = min(max(hpn, hpn_min), hpn_max)*h if wpn else 0 # CE Height(%), if width != 0
            ln = min(max(lpn, lpn_min), lpn_max)*h if wpn else 0 # EBA Height, if width != 0

            rect = self.shapes['%s-rect' % l] # section
            line = self.get_element_by_title('%s-line' % l) # dotted line(EBA)

            revenue = self.get_element_by_title('%s10' % l) # REV.
            revenue_arrow = self.get_element_by_title('%s11-arrow' % l) # REV. arrow
            earnings = self.get_element_by_title('%s13' % l) #ABA
            ce_marg = self.get_element_by_title('%s15' % l) # C.E
            ce_marg_p = self.get_element_by_title('%s16' % l) # C.E Marg

            # if section exists, set icon under it(Advisory, Audit etc)
            if wpn:
                icon = self.clone_template('icon-%s' % self.get_cell('%s9' % l).lower())
                self.set_element_pos(icon, x+self.get_element_coords(icon)[1]-coords[1], None)
                separate_chart_group.append(icon)

            # If width of a section is between 7%-4%
            if wpn <= .07:
                # Remove all values, just display correct heights of margin boxes/lines 
                for e in [revenue, revenue_arrow, earnings, ce_marg, ce_marg_p]:
                    remove(e)
            else:
                # If Width of a Section is above 7% of total width 
                # Set all values
                for e in [revenue, earnings, ce_marg, ce_marg_p]:
                    self.set_element_text_size(e, font_size)
                # display all values
                self.set_element_pos(revenue, x, None) # set REV.
                self.set_element_text(revenue, format_money(self.get_cell('%s10' % l)/10.**6))

                revenue_arrow_shift = self.get_element_sizes(revenue)[0]-self.get_element_sizes(revenue_arrow)[0] # set REV. arrow
                self.set_element_pos(revenue_arrow, x+revenue_arrow_shift/2, None)
                self.set_element_text(revenue_arrow, self.format_percent(self.get_cell('%s11' % l), prec=0)+'%')

                earnings_value = self.get_cell('%s13' % l)/10.**6 # set EBA
                self.set_element_pos(earnings, x, y0-ln)
                self.set_element_text(earnings, format_money(earnings_value))

                ce_marg_val = self.get_cell('%s15' % l)/10.**6 # set C.E
                self.set_element_pos(ce_marg, x, None)
                self.set_element_text(ce_marg, format_money(ce_marg_val))

                ce_marg_p_val = self.get_cell('%s16' % l) # set C.E Marg
                self.set_element_pos(ce_marg_p, x, None)
                self.set_element_text(ce_marg_p, self.format_percent(ce_marg_p_val, prec=1)+'%')

                marg_arrow_val = self.get_cell('%s17' % l) # set C.E Marg arrow
                marg_arrow = self.clone_template('marg-arrow-up' if marg_arrow_val >= 0 else 'marg-arrow-down')
                shift = self.get_element_sizes(ce_marg_p)[0]-self.get_element_sizes(marg_arrow)[0]
                self.set_element_text(marg_arrow, self.format_float(min(abs(marg_arrow_val), 999), prec=0)+'BPS')
                self.set_element_pos(marg_arrow, x+shift/2, None)
                separate_chart_group.append(marg_arrow)

            if wpn == 0:
                # if section width == 0, remove section
                remove(rect)
                remove(line)
            else: 
                # add section and EBA line
                self.set_element_pos(rect, x, y0-hn)
                self.set_element_size(rect, wn, hn)

                self.set_element_pos(line, x, y0-ln)
                self.set_element_size(line, wn, None)

            x += wn

            if l != 'E':
                s = self.get_elements_by_title('%s-sep' % l)[0]
                if wn == 0:
                    remove(s)
                else:
                    self.set_element_pos(s, x, None)

    # Revenue & Earnings – Actual Variance to Plan
    def fill_revenue_and_earnings(self):
        separate_chart_group = self.E('separate-revenue-and-earnings-variance')

        x0 = self.get_element_coords(self.get_element_by_title('rev-earn-min'))[0]
        x1 = self.get_element_coords(self.get_element_by_title('rev-earn-max'))[0]
        w = x1-x0 # chart width

        # Scale is Constant (-5% to 15%)
        minp = -.05 # x Axis Min
        maxp = .15 # x Axis Max
        step = .005 # tick step

        positions = defaultdict(dict)

        def set_tick(i, below, stack_level, min_stacked, max_stacked, total_stacked):
            name = '%s%s' % i
            tick = self.get_element_by_title('icon-'+name) 
            color = {
                'B': '898B8D', # Advisory
                'C': '1E2556', # Audit
                'D': '39A0DA', # Consulting
                'E': '43913E', # Tax
            }.get(i[0])

            show_arrow = False
            v = self.get_cell(name)
            tick_w, tick_h = self.get_element_sizes(tick)
            if v < minp:
                # If Value is below -5% 
                pos_v = minp-step # Display Chiclet at -5.5%
                show_arrow = True
            elif v > maxp:
                # If value is above 15% 
                pos_v = maxp+step # Display Chiclet at 15.5%
                show_arrow = True
            else:
                pos_v = v

            x = x0 + (pos_v-minp)/(maxp-minp)*w # normalized tick position
            xend = x+tick_w
            self.set_element_pos(tick, x-tick_w/2, None)
            level = 0
            # set labels position
            for (start, end), l in positions[below].iteritems():
                if start <= xend and x <= end:
                    level = max(level, l+1)
            positions[below][(x, xend)] = level

            text = '$'+self.format_float(self.get_cell('%s%s' % (i[0], i[1]+1))/10.**6, prec=0) 

            # 
            value_e = self.clone_template('rev-earn-label')
            value_e_w, value_e_h = self.get_element_sizes(value_e) # label width, height

            self.set_element_text(value_e, text)
            if color: # set labels color
                self.set_element_text_color(value_e, color)
            level_shift = 0 if i[0] == 'F' else (1 if below else -1)*level*value_e_h # height shift if stacked
            self.set_element_pos(
                value_e,
                x-value_e_w/2,
                (self.get_element_coords(tick)[3]
                    if below else (self.get_element_coords(tick)[1]-value_e_h))+level_shift
            )
            separate_chart_group.append(value_e)

            if show_arrow:
                shift = (1 if below else -1)*tick_h*stack_level if min_stacked or max_stacked else 0 # arrow shift
                self.mod_element_pos(tick, None, shift)
                arrow_e = self.clone_template('revenue-arrow-%s' % ('down' if v < 0 else 'up'))
                if min_stacked or max_stacked:
                    arrow_y = self.get_element_coords(tick)[3]-self.get_element_sizes(arrow_e)[1]+tick_h/2
                    self.set_element_pos(
                        value_e,
                        None,
                        self.get_element_coords(tick)[3]-self.get_element_sizes(value_e)[1]/2-tick_h/2
                    )
                else:
                    arrow_y = (
                        self.get_element_coords(tick)[1]-self.get_element_sizes(arrow_e)[1]
                        if below else
                        self.get_element_coords(tick)[3]
                    )
                if min_stacked:
                    arrow_x = x+tick_w/2
                    self.set_element_pos(value_e, x-tick_w/2-self.get_element_sizes(value_e)[0], None)
                elif max_stacked:
                    arrow_x = x-tick_w/2-self.get_element_sizes(arrow_e)[0]
                    self.set_element_pos(value_e, x+tick_w/2, None)
                else:
                    arrow_x = x-self.get_element_sizes(arrow_e)[0]/2
                self.set_element_pos(
                    arrow_e,
                    arrow_x,
                    arrow_y,
                )
                self.replace_pic_color(arrow_e, color or '92D050') # set color according to function or blue
                separate_chart_group.append(arrow_e)
                self.set_element_text(arrow_e, self.format_percent(abs(v), prec=1)+'%')

        for n in [22, 24]:
            cols = 'BCDEF' # exel cols
            cells = ['%s%s' % (c, n) for c in cols] # excel cells
            values = [self.get_cell(c) for c in cells] # cell's values
            min_stacked = sum(v < minp for v in values[:-1]) 
            max_stacked = sum(v > maxp for v in values[:-1])
            min_level = -1 # label level
            max_level = -1
            for col, v in zip(cols, values):
                level = 0 # set labels level
                if v < minp:
                    level = min_level = min_level+1
                if v > maxp:
                    level = max_level = max_level+1
                set_tick(
                    (col, n),
                    n == 24,
                    level,
                    min_stacked > 1 and v < minp,
                    max_stacked > 1 and v > maxp,
                    min_stacked if min_stacked > 1 and v < minp else max_stacked,
                )

    # Enabling Areas & Parent Costs ok
    def fill_enabling_areas(self):
        max_h = self.get_element_sizes(self.get_element_by_title('B41'))[1] # Largest value will be Max  height
        y1 = self.get_element_coords(self.get_element_by_title('enabling-areas-min'))[3]

        for c in alpha_range('B', 'K'):
            current_e = self.get_element_by_title('%s41' % c) # Height (% of Max). Current Amt
            plan_e = self.get_element_by_title('%s43' % c) # Height (% of Max). Plan
            prior_e = self.get_element_by_title('%s45' % c) # Height of Max. Prior

            current, plan, prior = [max(self.get_cell('%s%s' % (c, r)), 0) for r in [41, 43, 45]] # get Current Amt, Plan and Prior values

            # set Current Amt 
            self.set_element_size(current_e, None, max_h*current)
            self.set_element_pos(current_e, None, y1-max_h*current)

            # set Plan
            self.set_element_pos(
                plan_e,
                None, y1-max_h*plan-self.get_element_sizes(plan_e)[1]/2
            )   

            # set Prior
            self.set_element_pos(
                prior_e,
                None, y1-max_h*prior-self.get_element_sizes(prior_e)[1]/2
            )
            # set Current Amount label
            label_e = self.get_element_by_title('%s40' % c)
            self.set_element_pos(label_e, None, y1-max(current, prior, plan)*max_h-self.get_element_sizes(label_e)[1])

            # labels under sections(Parent, IT, Talent etc)
            self.set_element_text_lines(
                self.get_element_by_title('%s39' % c),
                [x[:9] for x in self.get_cell('%s39' % c).split()]
            )

    def fill_values(self):
        super(Dashboard11, self).fill_values()

        self.fill_key_metrics()
        self.fill_blue_rect()
        self.fill_revenue_and_earnings()
        self.fill_enabling_areas()


if __name__ == '__main__':
    CMDHandler(
        Dashboard11,
    )
